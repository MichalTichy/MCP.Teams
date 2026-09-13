using System.ComponentModel;
using ModelContextProtocol.Server;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class ChatTools(GraphClient graph)
{
    [McpServerTool(Name = "teams_list_chats"), Description("List chats for the connected Microsoft Teams user, newest first.")]
    public async Task<string> ListChats([Description("Maximum chats to return (1-50).")]
        int limit = 20) => GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get,
        $"me/chats?$top={GraphClient.Top(limit)}&$orderby=lastUpdatedDateTime desc"));

    [McpServerTool(Name = "teams_get_chat"), Description("Get one Teams chat by its ID.")]
    public async Task<string> GetChat([Description("Teams chat ID.")] string chatId) =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"chats/{GraphClient.Id(chatId)}"));

    [McpServerTool(Name = "teams_list_chat_messages"), Description("List recent messages in a Teams chat.")]
    public async Task<string> ListChatMessages(
        [Description("Teams chat ID.")] string chatId,
        [Description("Maximum messages to return (1-50).")]
        int limit = 20) => GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get,
        $"chats/{GraphClient.Id(chatId)}/messages?$top={GraphClient.Top(limit)}"));

    [McpServerTool(Name = "teams_get_chat_message"), Description("Get a specific Teams chat message.")]
    public async Task<string> GetChatMessage(
        [Description("Teams chat ID.")] string chatId,
        [Description("Teams message ID.")] string messageId) => GraphClient.Print(await graph.DelegatedAsync(
        HttpMethod.Get, $"chats/{GraphClient.Id(chatId)}/messages/{GraphClient.Id(messageId)}"));

    [McpServerTool(Name = "teams_list_chat_members"), Description("List members of a Teams chat.")]
    public async Task<string> ListChatMembers([Description("Teams chat ID.")] string chatId) =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"chats/{GraphClient.Id(chatId)}/members"));

    [McpServerTool(Name = "teams_send_chat_message"), Description("Send a plain-text message to an existing Teams chat. This causes an external side effect and must only be called from an approved execution plan.")]
    public async Task<string> SendChatMessage(
        [Description("Teams chat ID.")] string chatId,
        [Description("Plain-text message to send.")] string message)
    {
        GraphClient.Require(message, nameof(message));
        return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Post,
            $"chats/{GraphClient.Id(chatId)}/messages",
            new { body = new { contentType = "text", content = message } }));
    }
}
