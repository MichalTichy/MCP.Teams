using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Server;
using TeamsMcp.Chats.Models;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class ChatTools(GraphClient graph)
{
    [McpServerTool(Name = "teams_list_chats"), Description("List recent chats for the connected Microsoft Teams user, newest first. The separate selfChat target sends privately to the connected user: use chatId 'self' (48:notes), even when absent from the recent list. No self-chat link is needed.")]
    public async Task<string> ListChats([Description("Maximum chats to return (1-50).")]
        int limit = 20)
    {
        var response = await graph.DelegatedAsync(HttpMethod.Get, "me/chats?$top=50");
        var chats = response?["value"] is JsonArray value
            ? value.OfType<JsonObject>()
                .OrderByDescending(chat => ParseDate(chat["lastUpdatedDateTime"]))
                .Take(GraphClient.Top(limit))
                .Select(chat => chat.DeepClone())
                .ToArray()
            : [];

        return new JsonObject { ["value"] = new JsonArray(chats), ["selfChat"] = ChatTarget.DescribeSelf() }
            .ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool(Name = "teams_get_chat"), Description("Get a Teams chat by ID, or resolve 'self' (48:notes) as private delivery to the connected Microsoft user without asking for a link. Self-chat uses the connected profile because Graph's ordinary chat metadata endpoint may not support it.")]
    public async Task<string> GetChat([Description(ChatTarget.ParameterDescription)] string chatId)
    {
        if (!ChatTarget.IsSelf(chatId))
            return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"chats/{GraphClient.Id(ChatTarget.Resolve(chatId))}"));

        var me = await graph.DelegatedAsync(HttpMethod.Get, "me?$select=id,displayName,userPrincipalName");
        GraphClient.Require(me?["id"]?.GetValue<string>() ?? "", "Connected user ID");
        var target = ChatTarget.DescribeSelf();
        target["connectedUser"] = me;
        return GraphClient.Print(target);
    }

    [McpServerTool(Name = "teams_list_chat_messages"), Description("List recent messages in a Teams chat.")]
    public async Task<string> ListChatMessages(
        [Description(ChatTarget.ParameterDescription)] string chatId,
        [Description("Maximum messages to return (1-50).")]
        int limit = 20) => GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get,
        $"chats/{GraphClient.Id(ChatTarget.Resolve(chatId))}/messages?$top={GraphClient.Top(limit)}"));

    [McpServerTool(Name = "teams_get_chat_message"), Description("Get a specific Teams chat message.")]
    public async Task<string> GetChatMessage(
        [Description(ChatTarget.ParameterDescription)] string chatId,
        [Description("Teams message ID.")] string messageId) => GraphClient.Print(await graph.DelegatedAsync(
        HttpMethod.Get, $"chats/{GraphClient.Id(ChatTarget.Resolve(chatId))}/messages/{GraphClient.Id(messageId)}"));

    [McpServerTool(Name = "teams_list_chat_members"), Description("List members of a Teams chat.")]
    public async Task<string> ListChatMembers([Description(ChatTarget.ParameterDescription)] string chatId)
    {
        if (!ChatTarget.IsSelf(chatId))
            return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"chats/{GraphClient.Id(ChatTarget.Resolve(chatId))}/members"));

        var me = await graph.DelegatedAsync(HttpMethod.Get, "me?$select=id,displayName,userPrincipalName");
        var ownerId = GraphClient.Require(me?["id"]?.GetValue<string>() ?? "", "Connected user ID");
        return GraphClient.Print(new JsonObject
        {
            ["value"] = new JsonArray(new JsonObject
            {
                ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
                ["userId"] = ownerId,
                ["displayName"] = me?["displayName"]?.DeepClone()
            })
        });
    }

    [McpServerTool(Name = "teams_send_chat_message"), Description("Send a plain-text message to a Teams chat. Use chatId 'self' (48:notes) to send privately to the connected user without a chat link. This causes an external side effect and must only be called from an approved execution plan.")]
    public async Task<string> SendChatMessage(
        [Description(ChatTarget.ParameterDescription)] string chatId,
        [Description("Plain-text message to send.")] string message)
    {
        GraphClient.Require(message, nameof(message));
        return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Post,
            $"chats/{GraphClient.Id(ChatTarget.Resolve(chatId))}/messages",
            new { body = new { contentType = "text", content = message } }));
    }

    private static DateTimeOffset ParseDate(JsonNode? value) =>
        DateTimeOffset.TryParse(value?.GetValue<string>(), out var parsed)
            ? parsed
            : DateTimeOffset.MinValue;
}
