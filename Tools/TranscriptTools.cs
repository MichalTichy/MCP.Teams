using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Server;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class TranscriptTools(GraphClient graph)
{
    [McpServerTool(Name = "teams_list_recent_transcripts")]
    [Description("Lists metadata for recent Teams meeting and ad-hoc call transcripts organized by the connected account. It does not download transcript content.")]
    public async Task<string> ListRecentTranscriptsAsync(
        [Description("Inclusive UTC ISO 8601 start timestamp.")] string fromDateTime,
        [Description("Exclusive UTC ISO 8601 end timestamp.")] string toDateTime,
        [Description("Maximum number of transcript metadata records to return (1-100).")] int limit = 50,
        CancellationToken cancellationToken = default)
    {
        var from = NormalizeUtc(fromDateTime, nameof(fromDateTime));
        var to = NormalizeUtc(toDateTime, nameof(toDateTime));
        if (from >= to)
        {
            throw new ArgumentException("fromDateTime must be earlier than toDateTime.");
        }

        var me = await graph.DelegatedAsync(
            HttpMethod.Get,
            "me?$select=id,displayName",
            cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException("Microsoft Graph returned no connected user.");
        var userId = RequireNodeString(me, "id");
        var displayName = NodeString(me, "displayName") ?? "Connected Microsoft account";
        var requestedLimit = Math.Clamp(limit, 1, 100);

        var scheduledPath =
            $"users/{GraphClient.Id(userId)}/onlineMeetings/getAllTranscripts(" +
            $"meetingOrganizerUserId='{GraphClient.Id(userId)}'," +
            $"startDateTime={Uri.EscapeDataString(from.ToString("O"))}," +
            $"endDateTime={Uri.EscapeDataString(to.ToString("O"))})?$top={requestedLimit}";
        var adhocPath =
            $"users/{GraphClient.Id(userId)}/adhocCalls/getAllTranscripts(" +
            $"userId={GraphClient.Id(userId)}," +
            $"startDateTime={Uri.EscapeDataString(from.ToString("O"))}," +
            $"endDateTime={Uri.EscapeDataString(to.ToString("O"))})?$top={requestedLimit}";

        var scheduledTask = graph.ApplicationAsync(HttpMethod.Get, scheduledPath, cancellationToken: cancellationToken);
        var adhocTask = graph.ApplicationAsync(HttpMethod.Get, adhocPath, cancellationToken: cancellationToken);
        await Task.WhenAll(scheduledTask, adhocTask);

        var value = new JsonArray();
        AddTranscripts(value, await scheduledTask, "scheduledMeeting", userId, displayName);
        AddTranscripts(value, await adhocTask, "adhocCall", userId, displayName);

        var ordered = value
            .OfType<JsonObject>()
            .OrderByDescending(item => ParseDate(item["createdDateTime"]))
            .Take(requestedLimit)
            .Select(item => item.DeepClone())
            .ToArray();
        return new JsonObject { ["value"] = new JsonArray(ordered) }
            .ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool(Name = "teams_get_transcript")]
    [Description("Downloads one Teams transcript on demand. Supply meetingId for a scheduled meeting or callId for an ad-hoc call, but not both.")]
    public async Task<string> GetTranscriptAsync(
        [Description("Transcript ID returned by teams_list_recent_transcripts.")] string transcriptId,
        [Description("Scheduled online meeting ID, when the transcript belongs to a meeting.")] string? meetingId = null,
        [Description("Ad-hoc call ID, when the transcript belongs to an ad-hoc call.")] string? callId = null,
        CancellationToken cancellationToken = default)
    {
        var hasMeeting = !string.IsNullOrWhiteSpace(meetingId);
        var hasCall = !string.IsNullOrWhiteSpace(callId);
        if (hasMeeting == hasCall)
        {
            throw new ArgumentException("Supply exactly one of meetingId or callId.");
        }

        var me = await graph.DelegatedAsync(
            HttpMethod.Get,
            "me?$select=id",
            cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException("Microsoft Graph returned no connected user.");
        var userId = RequireNodeString(me, "id");
        var path = hasMeeting
            ? $"users/{GraphClient.Id(userId)}/onlineMeetings/{GraphClient.Id(meetingId!)}/transcripts/{GraphClient.Id(transcriptId)}/content"
            : $"users/{GraphClient.Id(userId)}/adhocCalls/{GraphClient.Id(callId!)}/transcripts/{GraphClient.Id(transcriptId)}/content";
        var content = await graph.ApplicationTextAsync(
            HttpMethod.Get,
            path,
            "text/vtt",
            cancellationToken);
        return new JsonObject
        {
            ["transcriptId"] = transcriptId.Trim(),
            ["meetingId"] = hasMeeting ? meetingId!.Trim() : null,
            ["callId"] = hasCall ? callId!.Trim() : null,
            ["format"] = "text/vtt",
            ["content"] = content
        }.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static void AddTranscripts(
        JsonArray target,
        JsonNode? response,
        string kind,
        string organizerId,
        string organizerName)
    {
        if (response?["value"] is not JsonArray items)
        {
            return;
        }

        foreach (var item in items.OfType<JsonObject>())
        {
            var id = NodeString(item, "id");
            var created = NodeString(item, "createdDateTime");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(created))
            {
                continue;
            }

            target.Add(new JsonObject
            {
                ["id"] = id,
                ["kind"] = kind,
                ["meetingId"] = NodeString(item, "meetingId"),
                ["callId"] = NodeString(item, "callId"),
                ["organizerId"] = organizerId,
                ["organizerName"] = organizerName,
                ["createdDateTime"] = created,
                ["endDateTime"] = NodeString(item, "endDateTime")
            });
        }
    }

    private static DateTimeOffset ParseDate(JsonNode? node) =>
        DateTimeOffset.TryParse(node?.GetValue<string>(), out var parsed)
            ? parsed
            : DateTimeOffset.MinValue;

    private static string RequireNodeString(JsonNode node, string propertyName) =>
        NodeString(node, propertyName)
        ?? throw new InvalidOperationException($"Microsoft Graph response has no '{propertyName}'.");

    private static string? NodeString(JsonNode node, string propertyName) =>
        node[propertyName]?.GetValue<string>();

    private static DateTimeOffset NormalizeUtc(string value, string parameterName)
    {
        if (!DateTimeOffset.TryParse(value, out var parsed))
        {
            throw new ArgumentException($"{parameterName} must be an ISO 8601 timestamp.", parameterName);
        }

        return parsed.ToUniversalTime();
    }
}
