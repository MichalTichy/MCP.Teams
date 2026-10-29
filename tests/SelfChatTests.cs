using System.Net;
using System.Text.Json.Nodes;
using ModelContextProtocol.Server;
using TeamsMcp.Tests.Support;
using TeamsMcp.Tools;

namespace TeamsMcp.Tests;

public sealed class SelfChatTests
{
    [Theory]
    [InlineData("self")]
    [InlineData("48:notes")]
    [InlineData(" SELF ")]
    public async Task Self_target_is_resolved_from_the_connected_account_without_chat_metadata(string target)
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = request => request.Path == "/v1.0/me"
            ? RecordingHttpHandler.Json(new { id = "owner-1", displayName = "Owner", userPrincipalName = "owner@example.com" })
            : throw new InvalidOperationException("Self-chat metadata is unavailable.");

        var result = JsonNode.Parse(await context.Chats.GetChat(target))!;

        Assert.Equal("48:notes", result["id"]?.GetValue<string>());
        Assert.True(result["isSelfChat"]?.GetValue<bool>());
        Assert.Equal("owner-1", result["connectedUser"]?["id"]?.GetValue<string>());
        Assert.Single(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Self_target_is_exposed_even_when_missing_from_recent_chats()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { value = new[] { new { id = "other-chat" } } });

        var result = JsonNode.Parse(await context.Chats.ListChats(1))!;

        Assert.Equal("other-chat", result["value"]?[0]?["id"]?.GetValue<string>());
        Assert.Single(result["value"]!.AsArray());
        Assert.Equal("48:notes", result["selfChat"]?["id"]?.GetValue<string>());
        Assert.All(context.GraphHttp.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
    }

    [Fact]
    public async Task Self_members_are_the_connected_user_without_using_the_chat_membership_endpoint()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = request => request.Path == "/v1.0/me"
            ? RecordingHttpHandler.Json(new { id = "owner-1", displayName = "Owner" })
            : throw new InvalidOperationException("Self-chat membership is unavailable.");

        var result = JsonNode.Parse(await context.Chats.ListChatMembers("self"))!;

        Assert.Equal("owner-1", Assert.Single(result["value"]!.AsArray())?["userId"]?.GetValue<string>());
        Assert.Single(context.GraphHttp.Requests);
    }

    [Theory]
    [InlineData("self")]
    [InlineData("48:notes")]
    public async Task Plain_text_delivery_posts_once_to_the_self_target(string target)
    {
        using var context = new ChatFilesTestContext();

        await context.Chats.SendChatMessage(target, "Private report");

        var sent = Assert.Single(context.GraphHttp.Requests);
        Assert.Equal(HttpMethod.Post, sent.Method);
        Assert.Equal("/v1.0/chats/48:notes/messages", Uri.UnescapeDataString(sent.Path));
        Assert.Equal("Private report", sent.Json?["body"]?["content"]?.GetValue<string>());
    }

    [Fact]
    public async Task Reading_self_messages_uses_the_same_target()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { value = Array.Empty<object>() });

        await context.Chats.ListChatMessages("self");
        await context.Chats.GetChatMessage("self", "message-1");

        Assert.Equal("/v1.0/chats/48:notes/messages", Uri.UnescapeDataString(context.GraphHttp.Requests[0].Path));
        Assert.Equal("/v1.0/chats/48:notes/messages/message-1", Uri.UnescapeDataString(context.GraphHttp.Requests[1].Path));
        Assert.All(context.GraphHttp.Requests, request => Assert.Equal(HttpMethod.Get, request.Method));
    }

    [Fact]
    public async Task Ordinary_chat_metadata_is_still_retrieved_from_graph()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { id = "chat-1" });

        var result = JsonNode.Parse(await context.Chats.GetChat("chat-1"));

        Assert.Equal("chat-1", result?["id"]?.GetValue<string>());
        Assert.Equal("/v1.0/chats/chat-1", Assert.Single(context.GraphHttp.Requests).Path);
    }

    [Fact]
    public async Task Self_lookup_does_not_hide_a_connection_failure()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { error = new { message = "Not connected" } }, HttpStatusCode.Unauthorized);

        await Assert.ThrowsAsync<TeamsMcp.Services.GraphRequestException>(() => context.Chats.GetChat("self"));
        Assert.Single(context.GraphHttp.Requests);
    }

    [Theory]
    [InlineData(nameof(ChatTools.GetChat))]
    [InlineData(nameof(ChatTools.SendChatMessage))]
    public void Existing_tool_schema_explains_delivery_to_the_connected_user(string method)
    {
        using var context = new ChatFilesTestContext();
        var tool = McpServerTool.Create(typeof(ChatTools).GetMethod(method)!, context.Chats);
        var description = tool.ProtocolTool.InputSchema.GetProperty("properties").GetProperty("chatId").GetProperty("description").GetString();

        Assert.Contains("self", description);
        Assert.Contains("connected Microsoft user", description);
        Assert.Contains("48:notes", tool.ProtocolTool.Description);
    }
}
