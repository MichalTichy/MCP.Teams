using System.ComponentModel;
using ModelContextProtocol.Server;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class ChannelTools(GraphClient graph)
{
    [McpServerTool(Name = "teams_list_joined_teams"), Description("List Microsoft Teams joined by the connected user.")]
    public async Task<string> ListJoinedTeams() =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, "me/joinedTeams"));

    [McpServerTool(Name = "teams_get_team"), Description("Get one Microsoft Team by ID.")]
    public async Task<string> GetTeam([Description("Team ID.")] string teamId) =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"teams/{GraphClient.Id(teamId)}"));

    [McpServerTool(Name = "teams_list_channels"), Description("List channels in a Microsoft Team.")]
    public async Task<string> ListChannels([Description("Team ID.")] string teamId) =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"teams/{GraphClient.Id(teamId)}/allChannels"));

    [McpServerTool(Name = "teams_get_channel"), Description("Get one Teams channel.")]
    public async Task<string> GetChannel(
        [Description("Team ID.")] string teamId,
        [Description("Channel ID.")] string channelId) => GraphClient.Print(await graph.DelegatedAsync(
        HttpMethod.Get, $"teams/{GraphClient.Id(teamId)}/channels/{GraphClient.Id(channelId)}"));

    [McpServerTool(Name = "teams_list_channel_messages"), Description("List root messages and replies in a Teams channel.")]
    public async Task<string> ListChannelMessages(
        [Description("Team ID.")] string teamId,
        [Description("Channel ID.")] string channelId,
        [Description("Maximum root messages to return (1-50).")]
        int limit = 20) => GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get,
        $"teams/{GraphClient.Id(teamId)}/channels/{GraphClient.Id(channelId)}/messages?$top={GraphClient.Top(limit)}&$expand=replies"));

    [McpServerTool(Name = "teams_send_channel_message"), Description("Post a plain-text root message to a Teams channel. This causes an external side effect and must only be called from an approved execution plan.")]
    public async Task<string> SendChannelMessage(
        [Description("Team ID.")] string teamId,
        [Description("Channel ID.")] string channelId,
        [Description("Plain-text message to post.")] string message)
    {
        GraphClient.Require(message, nameof(message));
        return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Post,
            $"teams/{GraphClient.Id(teamId)}/channels/{GraphClient.Id(channelId)}/messages",
            new { body = new { contentType = "text", content = message } }));
    }

    [McpServerTool(Name = "teams_reply_to_channel_message"), Description("Reply in the thread of an existing Teams channel message. This causes an external side effect and must only be called from an approved execution plan.")]
    public async Task<string> ReplyToChannelMessage(
        [Description("Team ID.")] string teamId,
        [Description("Channel ID.")] string channelId,
        [Description("Root channel message ID.")] string messageId,
        [Description("Plain-text reply to post.")] string message)
    {
        GraphClient.Require(message, nameof(message));
        return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Post,
            $"teams/{GraphClient.Id(teamId)}/channels/{GraphClient.Id(channelId)}/messages/{GraphClient.Id(messageId)}/replies",
            new { body = new { contentType = "text", content = message } }));
    }
}
