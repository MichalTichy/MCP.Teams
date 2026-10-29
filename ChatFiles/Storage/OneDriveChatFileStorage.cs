using System.Net;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.Chats.Models;
using TeamsMcp.Configuration;
using TeamsMcp.Services;

namespace TeamsMcp.ChatFiles.Storage;

public sealed class OneDriveChatFileStorage(GraphClient graph, IOptions<TeamsOptions> options)
{
    private const string FolderName = "Microsoft Teams Chat Files";

    public async Task<IReadOnlyList<JsonObject>> GetRecipientsAsync(string chatId, CancellationToken cancellationToken)
    {
        var me = await graph.DelegatedAsync(HttpMethod.Get, "me?$select=id", cancellationToken: cancellationToken);
        var ownerId = GraphClient.Require(me?["id"]?.GetValue<string>() ?? "", "Connected user ID");
        // The fixed self target belongs to the delegated account. Graph's regular
        // chat membership endpoint is not required or reliably available for it.
        if (ChatTarget.IsSelf(chatId)) return [];
        var recipients = new Dictionary<string, JsonObject>(StringComparer.OrdinalIgnoreCase);
        var ownerIsMember = false;
        string? path = $"chats/{GraphClient.Id(chatId)}/members";
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (path is not null)
        {
            if (!visited.Add(path) || visited.Count > 100) throw new InvalidOperationException("Chat membership pagination did not finish.");
            var page = await graph.DelegatedAsync(HttpMethod.Get, path, cancellationToken: cancellationToken);
            if (page?["value"] is not JsonArray members) throw new InvalidOperationException("Chat membership could not be verified.");
            foreach (var member in members)
            {
                if (member?["@odata.type"]?.GetValue<string>() != "#microsoft.graph.aadUserConversationMember")
                    throw new InvalidOperationException("Every chat member must be a resolvable Microsoft account before files can be shared.");
                var userId = member["userId"]?.GetValue<string>();
                var tenantId = member["tenantId"]?.GetValue<string>();
                if (string.Equals(userId, ownerId, StringComparison.OrdinalIgnoreCase) &&
                    (tenantId is null || string.Equals(tenantId, options.Value.TenantId, StringComparison.OrdinalIgnoreCase)))
                {
                    ownerIsMember = true;
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(userId) && string.Equals(tenantId, options.Value.TenantId, StringComparison.OrdinalIgnoreCase))
                    recipients["id:" + userId] = new JsonObject { ["objectId"] = userId };
                else if (member["email"]?.GetValue<string>() is { Length: > 0 } email)
                    recipients["email:" + email] = new JsonObject { ["email"] = email };
                else throw new InvalidOperationException("A chat member has no verifiable identity for file sharing.");
            }
            path = GetNextPagePath(page["@odata.nextLink"]?.GetValue<string>());
        }
        if (!ownerIsMember) throw new InvalidOperationException("The connected account is not a verified member of this chat.");
        return recipients.Values.ToArray();
    }

    public async Task<ChatFileReference> UploadAndShareAsync(
        ChatFileUpload file, byte[] bytes, IReadOnlyList<JsonObject> recipients,
        List<UploadedChatFile> uploadedFiles, CancellationToken cancellationToken)
    {
        var folderId = await GetFolderIdAsync(cancellationToken);
        var extension = Path.GetExtension(file.Name);
        var uniqueName = Path.GetFileNameWithoutExtension(file.Name) + "_" + Guid.NewGuid().ToString("N") + extension;
        var uploaded = await graph.UploadAsync(
            $"me/drive/items/{GraphClient.Id(folderId)}:/{Uri.EscapeDataString(uniqueName)}:/content",
            bytes, file.ContentType, cancellationToken);
        var driveItemId = GraphClient.Require(uploaded?["id"]?.GetValue<string>() ?? "", "Uploaded drive item ID");
        var index = uploadedFiles.Count;
        uploadedFiles.Add(new UploadedChatFile(driveItemId, uniqueName, uploaded?["webUrl"]?.GetValue<string>(), false));
        var itemPath = $"me/drive/items/{GraphClient.Id(driveItemId)}";
        var metadata = await graph.DelegatedAsync(HttpMethod.Get,
            itemPath + "?$select=id,name,webUrl,webDavUrl,eTag,sharepointIds", cancellationToken: cancellationToken);
        var guidText = metadata?["sharepointIds"]?["listItemUniqueId"]?.GetValue<string>() ??
            Regex.Match(metadata?["eTag"]?.GetValue<string>() ?? "", @"\{([0-9a-fA-F-]{36})\}").Groups[1].Value;
        if (!Guid.TryParse(guidText, out var fileId) || fileId == Guid.Empty)
            throw new InvalidOperationException("OneDrive did not return a SharePoint file GUID.");
        var webDavUrl = metadata?["webDavUrl"]?.GetValue<string>();
        if (!Uri.TryCreate(webDavUrl, UriKind.Absolute, out var fileUrl) || fileUrl.Scheme != Uri.UriSchemeHttps ||
            fileUrl.UserInfo.Length != 0 || fileUrl.Fragment.Length != 0)
            throw new InvalidOperationException("OneDrive did not return a valid HTTPS webDavUrl.");
        if (recipients.Count > 0)
        {
            var shared = await graph.DelegatedAsync(HttpMethod.Post, itemPath + "/invite", new
            {
                recipients, requireSignIn = true, sendInvitation = false, roles = new[] { "read" }
            }, cancellationToken);
            if (shared?["value"] is not JsonArray permissions || permissions.Count == 0 ||
                permissions.Any(permission => permission?["error"] is not null ||
                    string.IsNullOrEmpty(permission?["id"]?.GetValue<string>())))
                throw new InvalidOperationException("OneDrive did not confirm sharing with every chat recipient.");
            // Graph may return 207 with per-recipient errors; require all requested identities in successful permissions.
            foreach (var recipient in recipients)
            {
                if (!permissions.Any(permission => GrantsAccess(permission, recipient)))
                    throw new InvalidOperationException("OneDrive did not confirm access for every chat recipient.");
            }
        }
        uploadedFiles[index] = uploadedFiles[index] with { Shared = true };
        return new ChatFileReference { Id = fileId.ToString(), ContentUrl = webDavUrl!, Name = metadata?["name"]?.GetValue<string>() ?? uniqueName };
    }

    private static bool GrantsAccess(JsonNode? permission, JsonObject recipient)
    {
        if (permission?["roles"] is not JsonArray roles ||
            !roles.Any(role => role?.GetValue<string>() is "read" or "write" or "owner")) return false;
        var identities = new List<JsonNode?> { permission["grantedToV2"], permission["grantedTo"] };
        if (permission["grantedToIdentitiesV2"] is JsonArray many) identities.AddRange(many);
        if (permission["grantedToIdentities"] is JsonArray legacy) identities.AddRange(legacy);
        var objectId = recipient["objectId"]?.GetValue<string>();
        var email = recipient["email"]?.GetValue<string>();
        return objectId is not null
            ? identities.Any(identity => string.Equals(identity?["user"]?["id"]?.GetValue<string>(), objectId, StringComparison.OrdinalIgnoreCase))
            : email is not null && (string.Equals(permission["invitation"]?["email"]?.GetValue<string>(), email, StringComparison.OrdinalIgnoreCase) ||
                identities.Any(identity => string.Equals(identity?["user"]?["email"]?.GetValue<string>(), email, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(identity?["siteUser"]?["email"]?.GetValue<string>(), email, StringComparison.OrdinalIgnoreCase)));
    }

    private async Task<string> GetFolderIdAsync(CancellationToken cancellationToken)
    {
        var path = "me/drive/root:/" + Uri.EscapeDataString(FolderName);
        JsonNode? folder;
        try { folder = await graph.DelegatedAsync(HttpMethod.Get, path, cancellationToken: cancellationToken); }
        catch (GraphRequestException exception) when (exception.StatusCode == HttpStatusCode.NotFound)
        {
            try
            {
                folder = await graph.DelegatedAsync(HttpMethod.Post, "me/drive/root/children",
                    new JsonObject { ["name"] = FolderName, ["folder"] = new JsonObject(), ["@microsoft.graph.conflictBehavior"] = "fail" }, cancellationToken);
            }
            catch (GraphRequestException conflict) when (conflict.StatusCode == HttpStatusCode.Conflict)
            {
                folder = await graph.DelegatedAsync(HttpMethod.Get, path, cancellationToken: cancellationToken);
            }
        }
        if (folder?["folder"] is null) throw new InvalidOperationException("The Teams upload destination is not a OneDrive folder.");
        return GraphClient.Require(folder["id"]?.GetValue<string>() ?? "", "OneDrive folder ID");
    }

    private string? GetNextPagePath(string? nextLink)
    {
        if (nextLink is null) return null;
        var graphBase = new Uri(options.Value.GraphBaseUrl.TrimEnd('/') + "/");
        if (!Uri.TryCreate(nextLink, UriKind.Absolute, out var next) || !graphBase.IsBaseOf(next) || next.UserInfo.Length != 0)
            throw new InvalidOperationException("Graph returned an invalid membership pagination URL.");
        return next.AbsoluteUri;
    }
}
