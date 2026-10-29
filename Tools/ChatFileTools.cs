using System.ComponentModel;
using System.Text.Json;
using MediatR;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.ChatFiles.Usecases;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class ChatFileTools(ISender sender)
{
    [McpServerTool(Name = "teams_send_chat_message_with_attachments"), Description(
        "Send one message to an existing Teams chat with references to existing OneDrive/SharePoint files. Files must already be uploaded and shared with every intended recipient by a separate storage tool. " +
        "This tool only posts the Teams message; it does not upload, download, modify, or share files or verify recipients' file permissions. Requires an approved execution plan. Never retry automatically after an unknown send outcome.")]
    public async Task<CallToolResult> SendChatMessageWithAttachments(
        [Description("Existing Teams chat ID, including a self-chat if appropriate.")] string chatId,
        [Description("Plain-text message accompanying the files; HTML is escaped.")] string message,
        [Description("1-10 existing file references returned by the separate storage tool: id (SharePoint file GUID), contentUrl (HTTPS webDavUrl), and name (filename). File upload and recipient access must be completed before sending. Do not supply bytes, base64, Gateway artifact URLs, or local paths.")]
        List<ChatFileReference> attachments,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var result = await sender.Send(new SendChatFilesUsecase(chatId, message, attachments), cancellationToken);
            return new CallToolResult { Content = [new TextContentBlock { Text = GraphClient.Print(result) }] };
        }
        catch (ChatFileDeliveryException exception)
        {
            var details = JsonSerializer.Serialize(new
            {
                error = exception.Message,
                cause = exception.InnerException?.Message,
                messageOutcome = exception.MessageOutcome,
                attachments = exception.Attachments
            }, new JsonSerializerOptions(JsonSerializerDefaults.Web));
            return new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = details }] };
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or HttpRequestException)
        {
            return new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = exception.Message }] };
        }
    }
}
