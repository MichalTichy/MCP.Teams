using System.Text.Encodings.Web;
using System.Text.Json.Nodes;
using MediatR;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.ChatFiles.Notifications;
using TeamsMcp.ChatFiles.Storage;
using TeamsMcp.Services;

namespace TeamsMcp.ChatFiles.Usecases;

public sealed class SendChatFilesUsecaseHandler(GraphClient graph, IPublisher publisher,
    ChatFileSourceReader sourceReader, OneDriveChatFileStorage storage)
    : IRequestHandler<SendChatFilesUsecase, JsonNode?>
{
    public async Task<JsonNode?> Handle(SendChatFilesUsecase request, CancellationToken cancellationToken)
    {
        var chatId = GraphClient.Id(request.ChatId);
        GraphClient.Require(request.Message, nameof(request.Message));
        ArgumentNullException.ThrowIfNull(request.Attachments);
        var uploads = request.Files ?? [];
        if (request.Attachments.Count + uploads.Count is < 1 or > 10)
            throw new ArgumentException("Supply between 1 and 10 files or existing file references in total.");

        var ids = new HashSet<Guid>();
        var references = new List<ChatFileReference>();
        foreach (var attachment in request.Attachments)
        {
            if (attachment is null || !Guid.TryParse(attachment.Id, out var id) || id == Guid.Empty)
                throw new ArgumentException("Every attachment needs the SharePoint file GUID supplied by the storage tool.");
            if (!ids.Add(id))
                throw new ArgumentException("The same file GUID cannot be attached twice in one message.");
            if (!Uri.TryCreate(attachment.ContentUrl, UriKind.Absolute, out var url) ||
                url.Scheme != Uri.UriSchemeHttps || url.UserInfo.Length != 0 ||
                url.Fragment.Length != 0 || string.IsNullOrEmpty(url.Host))
                throw new ArgumentException("contentUrl must be the existing file's HTTPS webDavUrl, without credentials or a fragment.");
            if (string.IsNullOrWhiteSpace(attachment.Name) || attachment.Name.Length > 255 || attachment.Name.Any(char.IsControl))
                throw new ArgumentException("Every attachment needs a display filename without control characters, at most 255 characters long.");
            references.Add(attachment with { Id = id.ToString() });
        }

        foreach (var upload in uploads) sourceReader.Validate(upload);
        var messageOutcome = "notSent";
        var uploadedFiles = new List<UploadedChatFile>();
        try
        {
            if (uploads.Count > 0)
            {
                // Read every source before making storage changes, so invalid later inputs cannot leave partial uploads.
                var sources = new List<byte[]>();
                long totalBytes = 0;
                foreach (var upload in uploads)
                {
                    var bytes = await sourceReader.ReadAsync(upload, cancellationToken);
                    totalBytes += bytes.LongLength;
                    if (totalBytes > 50 * 1024 * 1024) throw new ArgumentException("Combined file contents must not exceed 50 MiB.");
                    sources.Add(bytes);
                }
                var recipients = await storage.GetRecipientsAsync(request.ChatId, cancellationToken);
                for (var index = 0; index < uploads.Count; index++)
                {
                    var reference = await storage.UploadAndShareAsync(uploads[index], sources[index], recipients, uploadedFiles, cancellationToken);
                    if (!ids.Add(Guid.Parse(reference.Id))) throw new InvalidOperationException("OneDrive returned a duplicate attachment GUID.");
                    references.Add(reference);
                }
            }
            var content = string.Join("<br>", request.Message.Replace("\r\n", "\n").Replace('\r', '\n')
                .Split('\n').Select(HtmlEncoder.Default.Encode));
            content += string.Concat(references.Select(file => $"<attachment id=\"{file.Id}\"></attachment>"));
            messageOutcome = "unknown";
            var sent = await graph.DelegatedAsync(HttpMethod.Post, $"chats/{chatId}/messages", new
            {
                body = new { contentType = "html", content },
                attachments = references.Select(file => new
                {
                    id = file.Id,
                    contentType = "reference",
                    contentUrl = file.ContentUrl,
                    name = file.Name
                }).ToArray()
            }, cancellationToken);
            var messageId = GraphClient.Require(sent?["id"]?.GetValue<string>() ?? "", "Microsoft Graph message id");
            messageOutcome = "sent";
            await publisher.Publish(new ChatFilesMessageSentNotification(request.ChatId, messageId, references.ToArray()), cancellationToken);
            return sent;
        }
        catch (Exception exception)
        {
            // A missing receipt or transport failure cannot establish whether Graph accepted the send.
            throw new ChatFileDeliveryException(references.ToArray(), messageOutcome, exception, uploadedFiles.ToArray());
        }
    }
}
