using System.ComponentModel;
using System.Text.Json;
using MediatR;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.ChatFiles.Usecases;
using TeamsMcp.Chats.Models;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class ChatFileTools(ISender sender)
{
    [McpServerTool(Name = "teams_send_chat_message_with_attachments"), Description(
        "Send one Teams message with files in a single call. Supply files to upload to the connected user's OneDrive and share with every chat member automatically, or attachments for already shared file references, or both (1-10 total). " +
        "Use chatId 'self' (48:notes) to send files privately to the connected user; no chat link or sharing invitation is needed. " +
        "Requires an approved execution plan and Files.ReadWrite for uploads. Never retry automatically after a partial failure or unknown send outcome.")]
    public async Task<CallToolResult> SendChatMessageWithAttachments(
        [Description(ChatTarget.ParameterDescription)] string chatId,
        [Description("Plain-text message accompanying the files; HTML is escaped.")] string message,
        [Description("Optional existing file references: id (SharePoint GUID), contentUrl (HTTPS webDavUrl), name. These files must already be shared with recipients. For new files use files instead.")]
        List<ChatFileReference>? attachments = null,
        [Description("Optional new files: name, contentType, and either agentFileUrl or base64Content. Gateway accepts ./artifacts/... as agentFileUrl for generated files. Files are uploaded and shared before sending.")]
        List<ChatFileUpload>? files = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var result = await sender.Send(new SendChatFilesUsecase(chatId, message, attachments ?? [], files), cancellationToken);
            return new CallToolResult { Content = [new TextContentBlock { Text = GraphClient.Print(result) }] };
        }
        catch (ChatFileDeliveryException exception)
        {
            var details = JsonSerializer.Serialize(new
            {
                error = exception.Message,
                cause = exception.InnerException?.Message,
                messageOutcome = exception.MessageOutcome,
                attachments = exception.Attachments,
                uploadedFiles = exception.UploadedFiles
            }, new JsonSerializerOptions(JsonSerializerDefaults.Web));
            return new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = details }] };
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or HttpRequestException)
        {
            return new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = exception.Message }] };
        }
    }
}
