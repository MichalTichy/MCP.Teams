using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.Tests.Support;

namespace TeamsMcp.Tests;

public sealed class ChatFileUploadTests
{
    private const string FileGuid = "153fa47d-18c9-4179-be08-9879815a9f90";
    private const string ArtifactUrl = "http://gateway:8080/_internal/mcp/artifacts/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    [Fact]
    public async Task Inline_file_is_uploaded_shared_and_sent_in_one_tool_call()
    {
        using var context = CreateContext();
        var bytes = "real workbook bytes"u8.ToArray();
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline(bytes)]);

        Assert.NotEqual(true, result.IsError);
        var upload = Assert.Single(context.GraphHttp.Requests, request => request.Method == HttpMethod.Put);
        Assert.Equal(bytes, upload.Bytes);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", upload.ContentType);
        Assert.Equal("Bearer test-delegated", upload.Authorization);
        var invitation = Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/invite"));
        Assert.Equal("recipient-1", invitation.Json?["recipients"]?[0]?["objectId"]?.GetValue<string>());
        Assert.Equal("read", invitation.Json?["roles"]?[0]?.GetValue<string>());
        Assert.False(invitation.Json?["sendInvitation"]?.GetValue<bool>());
        var sent = Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
        Assert.Equal(FileGuid, sent.Json?["attachments"]?[0]?["id"]?.GetValue<string>());
        Assert.Contains("webDavUrl", context.GraphHttp.Requests.Single(request => request.Url.Query.Contains("webDavUrl")).Url.Query);
        Assert.Single(context.Notifications);
        Assert.Empty(context.FileHttp.Requests);
    }

    [Fact]
    public async Task Gateway_file_is_downloaded_without_graph_credentials_and_uploaded_unchanged()
    {
        using var context = CreateContext();
        var bytes = new byte[] { 80, 75, 3, 4, 0, 255 };
        context.FileHttp.Respond = _ => new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) };
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files:
            [new ChatFileUpload { Name = "report.xlsx", AgentFileUrl = ArtifactUrl }]);

        Assert.NotEqual(true, result.IsError);
        Assert.Null(Assert.Single(context.FileHttp.Requests).Authorization);
        Assert.Equal(bytes, Assert.Single(context.GraphHttp.Requests, request => request.Method == HttpMethod.Put).Bytes);
    }

    [Fact]
    public async Task Self_chat_needs_no_sharing_invitation()
    {
        using var context = CreateContext(selfChat: true);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        Assert.NotEqual(true, result.IsError);
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Path.EndsWith("/invite"));
        Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
    }

    [Theory]
    [InlineData("self")]
    [InlineData("48:notes")]
    public async Task Native_self_chat_upload_does_not_require_membership_lookup_or_file_sharing(string target)
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/members") || request.Path.EndsWith("/invite")
            ? throw new InvalidOperationException("Self-chat has no ordinary membership endpoint or other recipients.")
            : normal(request);

        var result = await context.Tool.SendChatMessageWithAttachments(target, "Private report", files: [Inline()]);

        Assert.NotEqual(true, result.IsError);
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Path.EndsWith("/members") || request.Path.EndsWith("/invite"));
        var sent = Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
        Assert.Equal("/v1.0/chats/48:notes/messages", Uri.UnescapeDataString(sent.Path));
        Assert.Equal("48:notes", Assert.Single(context.Notifications).ChatId);
    }

    [Fact]
    public async Task Self_chat_upload_stops_before_storage_changes_when_connected_identity_is_missing()
    {
        using var context = CreateContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { });

        var result = await context.Tool.SendChatMessageWithAttachments("self", "Private report", files: [Inline()]);

        Assert.True(result.IsError);
        Assert.Equal(HttpMethod.Get, Assert.Single(context.GraphHttp.Requests).Method);
        Assert.Empty(context.Notifications);
    }

    [Fact]
    public async Task Later_membership_page_is_included_before_sharing()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/members")
            ? request.Url.Query.Length == 0
                ? RecordingHttpHandler.Json(new Dictionary<string, object>
                {
                    ["value"] = new[] { Member("owner-1") },
                    ["@odata.nextLink"] = "https://graph.example.com/v1.0/chats/chat-1/members?skip=1"
                })
                : RecordingHttpHandler.Json(new { value = new[] { Member("recipient-1") } })
            : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        Assert.NotEqual(true, result.IsError);
        Assert.Equal(2, context.GraphHttp.Requests.Count(request => request.Path.EndsWith("/members")));
        Assert.Contains(context.GraphHttp.Requests, request => request.Path.EndsWith("/invite"));
    }

    [Fact]
    public async Task Missing_teams_folder_is_created_before_upload()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.Contains("root:")
            ? RecordingHttpHandler.Json(new { error = new { message = "Not found" } }, HttpStatusCode.NotFound)
            : request.Path.EndsWith("/root/children")
                ? RecordingHttpHandler.Json(new { id = "folder-1", folder = new { } }, HttpStatusCode.Created)
                : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        Assert.NotEqual(true, result.IsError);
        var create = Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/root/children"));
        Assert.Equal("Microsoft Teams Chat Files", create.Json?["name"]?.GetValue<string>());
        Assert.Equal("fail", create.Json?["@microsoft.graph.conflictBehavior"]?.GetValue<string>());
    }

    [Theory]
    [InlineData("denied")]
    [InlineData("partial")]
    [InlineData("wrong-recipient")]
    public async Task Unconfirmed_sharing_prevents_send_and_reports_uploaded_file(string failure)
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/invite")
            ? failure switch
            {
                "denied" => RecordingHttpHandler.Json(new { error = new { message = "Access denied" } }, HttpStatusCode.Forbidden),
                "partial" => RecordingHttpHandler.Json(new { value = new[] { new { error = new { message = "Cannot share" } } } }, HttpStatusCode.MultiStatus),
                _ => RecordingHttpHandler.Json(new { value = new[] { new { id = "permission-1", roles = new[] { "read" }, grantedToV2 = new { user = new { id = "another-user" } } } } })
            }
            : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        var error = Error(result);
        Assert.Equal("notSent", error["messageOutcome"]?.GetValue<string>());
        Assert.Equal("item-1", error["uploadedFiles"]?[0]?["driveItemId"]?.GetValue<string>());
        Assert.False(error["uploadedFiles"]?[0]?["shared"]?.GetValue<bool>());
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
        Assert.Empty(context.Notifications);
    }

    [Fact]
    public async Task Send_timeout_after_upload_does_not_repeat_upload_or_message()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/messages")
            ? throw new HttpRequestException("Connection lost") : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        var error = Error(result);
        Assert.Equal("unknown", error["messageOutcome"]?.GetValue<string>());
        Assert.True(error["uploadedFiles"]?[0]?["shared"]?.GetValue<bool>());
        Assert.Single(context.GraphHttp.Requests, request => request.Method == HttpMethod.Put);
        Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
    }

    [Theory]
    [InlineData("https://evil.example/_internal/mcp/artifacts/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")]
    [InlineData("http://gateway:8080/admin")]
    [InlineData("http://gateway:8080/_internal/mcp/artifacts/short")]
    [InlineData("http://user:password@gateway:8080/_internal/mcp/artifacts/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")]
    [InlineData("file:///workspace/artifacts/report.xlsx")]
    public async Task Invalid_download_targets_are_rejected_without_network_calls(string url)
    {
        using var context = CreateContext();
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files:
            [new ChatFileUpload { Name = "report.xlsx", AgentFileUrl = url }]);
        Assert.True(result.IsError);
        Assert.Empty(context.GraphHttp.Requests);
        Assert.Empty(context.FileHttp.Requests);
    }

    [Theory]
    [InlineData("../report.xlsx")]
    [InlineData("bad\r\nname.xlsx")]
    [InlineData("report?.xlsx")]
    public async Task Invalid_later_filename_is_rejected_before_any_download_or_upload(string name)
    {
        using var context = CreateContext();
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline(), Inline() with { Name = name }]);
        Assert.True(result.IsError);
        Assert.Empty(context.GraphHttp.Requests);
        Assert.Empty(context.FileHttp.Requests);
    }

    [Fact]
    public async Task Invalid_later_base64_prevents_all_graph_changes()
    {
        using var context = CreateContext();
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline(), Inline() with { Base64Content = "invalid!" }]);
        Assert.Equal("notSent", Error(result)["messageOutcome"]?.GetValue<string>());
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Missing_owner_membership_prevents_upload()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/members")
            ? RecordingHttpHandler.Json(new { value = new[] { Member("recipient-1") } }) : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        Assert.Equal("notSent", Error(result)["messageOutcome"]?.GetValue<string>());
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Method != HttpMethod.Get);
    }

    [Fact]
    public async Task Redirected_gateway_download_is_not_uploaded()
    {
        using var context = CreateContext();
        context.FileHttp.Respond = _ => new HttpResponseMessage(HttpStatusCode.Redirect)
            { Headers = { Location = new Uri("https://evil.example/file") } };
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files:
            [new ChatFileUpload { Name = "report.xlsx", AgentFileUrl = ArtifactUrl }]);
        Assert.Equal("notSent", Error(result)["messageOutcome"]?.GetValue<string>());
        Assert.Single(context.FileHttp.Requests);
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Oversized_gateway_download_is_rejected_from_headers_before_upload()
    {
        using var context = CreateContext();
        context.FileHttp.Respond = _ =>
        {
            var content = new ByteArrayContent([]);
            content.Headers.ContentLength = 26L * 1024 * 1024;
            return new HttpResponseMessage(HttpStatusCode.OK) { Content = content };
        };
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files:
            [new ChatFileUpload { Name = "report.xlsx", AgentFileUrl = ArtifactUrl }]);
        Assert.Equal("notSent", Error(result)["messageOutcome"]?.GetValue<string>());
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Multiple_uploads_produce_one_message_and_preserve_all_bytes()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        var metadataCount = 0;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/items/item-1")
            ? RecordingHttpHandler.Json(new
            {
                id = "item-1", name = "report.xlsx", webDavUrl = "https://tenant.sharepoint.com/Documents/report.xlsx",
                sharepointIds = new { listItemUniqueId = ++metadataCount == 1 ? FileGuid : "668f7fa8-8129-4de7-b32b-fe1b442e6ef1" }
            }) : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Reports", files:
            [Inline([1, 2, 3]), Inline([4, 5, 6]) with { Name = "invoice.pdf", ContentType = "application/pdf" }]);
        Assert.NotEqual(true, result.IsError);
        var uploads = context.GraphHttp.Requests.Where(request => request.Method == HttpMethod.Put).ToArray();
        Assert.Equal(2, uploads.Length);
        Assert.Equal(new byte[] { 1, 2, 3 }, uploads[0].Bytes);
        Assert.Equal(new byte[] { 4, 5, 6 }, uploads[1].Bytes);
        Assert.NotEqual(uploads[0].Path, uploads[1].Path);
        var sent = Assert.Single(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
        Assert.Equal(2, sent.Json?["attachments"]?.AsArray().Count);
    }

    [Fact]
    public async Task Failure_of_a_later_upload_reports_earlier_files_without_sending_a_message()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        var uploadCount = 0;
        context.GraphHttp.Respond = request => request.Method == HttpMethod.Put && ++uploadCount == 2
            ? RecordingHttpHandler.Json(new { error = new { message = "Quota exceeded" } }, HttpStatusCode.InsufficientStorage)
            : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Reports", files: [Inline(), Inline() with { Name = "other.xlsx" }]);
        var error = Error(result);
        Assert.Equal("notSent", error["messageOutcome"]?.GetValue<string>());
        Assert.Single(error["uploadedFiles"]!.AsArray());
        Assert.True(error["uploadedFiles"]?[0]?["shared"]?.GetValue<bool>());
        Assert.Equal(2, uploadCount);
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Path.EndsWith("/messages"));
    }

    [Fact]
    public async Task Hostile_membership_pagination_does_not_forward_graph_credentials_to_another_host()
    {
        using var context = CreateContext();
        var normal = context.GraphHttp.Respond!;
        context.GraphHttp.Respond = request => request.Path.EndsWith("/members")
            ? RecordingHttpHandler.Json(new Dictionary<string, object>
            {
                ["value"] = new[] { Member("owner-1") }, ["@odata.nextLink"] = "https://evil.example/members"
            }) : normal(request);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", files: [Inline()]);
        Assert.Equal("notSent", Error(result)["messageOutcome"]?.GetValue<string>());
        Assert.All(context.GraphHttp.Requests, request => Assert.Equal("graph.example.com", request.Url.Host));
        Assert.DoesNotContain(context.GraphHttp.Requests, request => request.Method != HttpMethod.Get);
    }

    private static ChatFileUpload Inline(byte[]? bytes = null) => new()
    {
        Name = "report.xlsx", ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        Base64Content = Convert.ToBase64String(bytes ?? Encoding.UTF8.GetBytes("report"))
    };

    private static JsonNode Error(CallToolResult result)
    {
        Assert.True(result.IsError);
        return JsonNode.Parse(Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text)!;
    }

    private static Dictionary<string, string> Member(string id) => new()
    {
        ["@odata.type"] = "#microsoft.graph.aadUserConversationMember", ["userId"] = id, ["tenantId"] = "tenant-1"
    };

    private static ChatFilesTestContext CreateContext(bool selfChat = false)
    {
        var context = new ChatFilesTestContext();
        string? storedName = null;
        context.GraphHttp.Respond = request =>
        {
            if (request.Path.EndsWith("/me")) return RecordingHttpHandler.Json(new { id = "owner-1" });
            if (request.Path.EndsWith("/members")) return RecordingHttpHandler.Json(new
                { value = selfChat ? new[] { Member("owner-1") } : new[] { Member("owner-1"), Member("recipient-1") } });
            if (request.Path.Contains("root:")) return RecordingHttpHandler.Json(new { id = "folder-1", folder = new { } });
            if (request.Method == HttpMethod.Put)
            {
                storedName = Uri.UnescapeDataString(request.Path.Split('/')[^2].TrimEnd(':'));
                return RecordingHttpHandler.Json(new { id = "item-1", webUrl = "https://tenant.sharepoint.com/file" }, HttpStatusCode.Created);
            }
            if (request.Path.EndsWith("/items/item-1")) return RecordingHttpHandler.Json(new
            {
                id = "item-1", name = storedName, webDavUrl = "https://tenant.sharepoint.com/Documents/report.xlsx",
                sharepointIds = new { listItemUniqueId = FileGuid }
            });
            if (request.Path.EndsWith("/invite")) return RecordingHttpHandler.Json(new
                { value = new[] { new { id = "permission-1", roles = new[] { "read" }, grantedToV2 = new { user = new { id = "recipient-1" } } } } });
            if (request.Path.EndsWith("/messages")) return RecordingHttpHandler.Json(new { id = "message-1" }, HttpStatusCode.Created);
            throw new InvalidOperationException("Unexpected test request: " + request.Url);
        };
        return context;
    }
}
