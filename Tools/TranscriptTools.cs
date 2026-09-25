using System.Collections.Concurrent;
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
    [Description("Lists metadata for recent Teams meeting and ad-hoc call transcripts accessible to the connected account, including invited calendar meetings. Independent discovery failures are returned without discarding successful results. It does not download transcript content.")]
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

        var transcripts = new ConcurrentBag<JsonObject>();
        var errors = new ConcurrentBag<string>();
        var skippedMeetings = new ConcurrentBag<string>();

        var scheduledPath =
            $"users/{GraphClient.Id(userId)}/onlineMeetings/getAllTranscripts(" +
            $"meetingOrganizerUserId='{GraphClient.Id(userId)}'," +
            $"startDateTime={Uri.EscapeDataString(from.ToString("O"))}," +
            $"endDateTime={Uri.EscapeDataString(to.ToString("O"))})?$top={requestedLimit}";
        var adhocPath =
            $"users/{GraphClient.Id(userId)}/adhocCalls/getAllTranscripts(" +
            $"userId='{GraphClient.Id(userId)}'," +
            $"startDateTime={Uri.EscapeDataString(from.ToString("O"))}," +
            $"endDateTime={Uri.EscapeDataString(to.ToString("O"))})?$top={requestedLimit}";

        await AddApplicationTranscriptsAsync(
            scheduledPath, "scheduledMeeting", "Organized scheduled meetings",
            userId, displayName, transcripts, errors, cancellationToken);
        await AddApplicationTranscriptsAsync(
            adhocPath, "adhocCall", "Ad-hoc calls",
            userId, displayName, transcripts, errors, cancellationToken);
        await AddInvitedMeetingTranscriptsAsync(
            from, to, requestedLimit, transcripts, errors, skippedMeetings, cancellationToken);

        var ordered = transcripts
            .GroupBy(TranscriptKey, StringComparer.Ordinal)
            .Select(group => group.First())
            .OrderByDescending(item => ParseDate(item["createdDateTime"]))
            .Take(requestedLimit)
            .Select(item => item.DeepClone())
            .ToArray();
        return new JsonObject
        {
            ["value"] = new JsonArray(ordered),
            ["errors"] = ToJsonArray(GroupMessages(errors)),
            ["skippedMeetings"] = ToJsonArray(GroupMessages(skippedMeetings))
        }
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
        var content = hasMeeting
            ? await DownloadMeetingTranscriptAsync(userId, meetingId!, transcriptId, cancellationToken)
            : await graph.ApplicationTextAsync(
                HttpMethod.Get,
                $"users/{GraphClient.Id(userId)}/adhocCalls/{GraphClient.Id(callId!)}/transcripts/{GraphClient.Id(transcriptId)}/content",
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

    private async Task<string> DownloadMeetingTranscriptAsync(
        string userId,
        string meetingId,
        string transcriptId,
        CancellationToken cancellationToken)
    {
        var suffix = $"onlineMeetings/{GraphClient.Id(meetingId)}/transcripts/{GraphClient.Id(transcriptId)}/content";
        try
        {
            return await graph.DelegatedTextAsync(
                HttpMethod.Get,
                $"me/{suffix}",
                "text/vtt",
                cancellationToken);
        }
        catch (InvalidOperationException delegatedFailure)
        {
            try
            {
                return await graph.ApplicationTextAsync(
                    HttpMethod.Get,
                    $"users/{GraphClient.Id(userId)}/{suffix}",
                    "text/vtt",
                    cancellationToken);
            }
            catch (InvalidOperationException applicationFailure)
            {
                throw new InvalidOperationException(
                    $"The transcript could not be downloaded with delegated or application access. " +
                    $"Delegated: {delegatedFailure.Message} Application: {applicationFailure.Message}",
                    applicationFailure);
            }
        }
    }

    private async Task AddApplicationTranscriptsAsync(
        string path,
        string kind,
        string source,
        string userId,
        string displayName,
        ConcurrentBag<JsonObject> transcripts,
        ConcurrentBag<string> errors,
        CancellationToken cancellationToken)
    {
        try
        {
            AddTranscripts(
                transcripts,
                await graph.ApplicationAsync(HttpMethod.Get, path, cancellationToken: cancellationToken),
                kind,
                userId,
                displayName);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            errors.Add($"{source}: {exception.Message}");
        }
    }

    private async Task AddInvitedMeetingTranscriptsAsync(
        DateTimeOffset from,
        DateTimeOffset to,
        int requestedLimit,
        ConcurrentBag<JsonObject> transcripts,
        ConcurrentBag<string> errors,
        ConcurrentBag<string> skippedMeetings,
        CancellationToken cancellationToken)
    {
        JsonNode? calendarResponse;
        try
        {
            var calendarFrom = from.AddDays(-7);
            var calendarPath = "me/calendarView?startDateTime="
                + Uri.EscapeDataString(calendarFrom.ToString("O"))
                + "&endDateTime=" + Uri.EscapeDataString(to.ToString("O"))
                + "&$select=onlineMeeting,organizer&$top=1000";
            calendarResponse = await graph.DelegatedAsync(
                HttpMethod.Get, calendarPath, cancellationToken: cancellationToken);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            errors.Add($"Invited meeting calendar discovery: {exception.Message}");
            return;
        }

        var candidates = Values(calendarResponse)
            .Select(item => new
            {
                JoinUrl = NodeString(item["onlineMeeting"], "joinUrl"),
                OrganizerName = NodeString(item["organizer"]?["emailAddress"], "name")
                    ?? "Meeting organizer"
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.JoinUrl))
            .DistinctBy(item => item.JoinUrl, StringComparer.Ordinal)
            .ToArray();

        await Parallel.ForEachAsync(candidates, new ParallelOptions
        {
            MaxDegreeOfParallelism = 4,
            CancellationToken = cancellationToken
        }, async (candidate, token) =>
        {
            try
            {
                var escapedJoinUrl = candidate.JoinUrl!.Replace("'", "''", StringComparison.Ordinal);
                var meetingPath = "me/onlineMeetings?$filter="
                    + Uri.EscapeDataString($"JoinWebUrl eq '{escapedJoinUrl}'");
                var meetingResponse = await graph.DelegatedAsync(
                    HttpMethod.Get, meetingPath, cancellationToken: token);
                foreach (var meeting in Values(meetingResponse))
                {
                    var meetingId = RequireNodeString(meeting, "id");
                    var organizerId = NodeString(meeting["participants"]?["organizer"]?["identity"]?["user"], "id");
                    var transcriptResponse = await graph.DelegatedAsync(
                        HttpMethod.Get,
                        $"me/onlineMeetings/{GraphClient.Id(meetingId)}/transcripts?$top={requestedLimit}",
                        cancellationToken: token);
                    AddTranscripts(
                        transcripts,
                        transcriptResponse,
                        "scheduledMeeting",
                        organizerId,
                        candidate.OrganizerName,
                        meetingId);
                }
            }
            catch (InvalidOperationException exception) when (IsInaccessibleMeeting(exception.Message))
            {
                skippedMeetings.Add(exception.Message);
            }
            catch (Exception exception) when (exception is not OperationCanceledException)
            {
                errors.Add($"Invited meeting: {exception.Message}");
            }
        });
    }

    private static void AddTranscripts(
        ConcurrentBag<JsonObject> target,
        JsonNode? response,
        string kind,
        string? organizerId,
        string organizerName,
        string? resolvedMeetingId = null)
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
                ["meetingId"] = resolvedMeetingId ?? NodeString(item, "meetingId"),
                ["callId"] = NodeString(item, "callId"),
                ["organizerId"] = organizerId,
                ["organizerName"] = organizerName,
                ["createdDateTime"] = created,
                ["endDateTime"] = NodeString(item, "endDateTime")
            });
        }
    }

    private static IEnumerable<JsonObject> Values(JsonNode? response) =>
        response?["value"] is JsonArray items
            ? items.OfType<JsonObject>()
            : [];

    private static bool IsInaccessibleMeeting(string message) =>
        (message.Contains("Microsoft Graph returned 403", StringComparison.Ordinal)
         && message.Contains("3003:", StringComparison.Ordinal))
        || (message.Contains("Microsoft Graph returned 404", StringComparison.Ordinal)
            && message.Contains("3004:", StringComparison.Ordinal));

    private static string TranscriptKey(JsonObject item) => string.Join('|',
        NodeString(item, "kind") ?? string.Empty,
        NodeString(item, "meetingId") ?? string.Empty,
        NodeString(item, "callId") ?? string.Empty,
        NodeString(item, "id") ?? string.Empty);

    private static IReadOnlyList<string> GroupMessages(IEnumerable<string> messages) => messages
        .GroupBy(message => message, StringComparer.Ordinal)
        .Select(group => group.Count() == 1 ? group.Key : $"{group.Count()} occurrences: {group.Key}")
        .Order(StringComparer.Ordinal)
        .ToArray();

    private static JsonArray ToJsonArray(IEnumerable<string> values) =>
        new(values.Select(value => JsonValue.Create(value)).ToArray());

    private static DateTimeOffset ParseDate(JsonNode? node) =>
        DateTimeOffset.TryParse(node?.GetValue<string>(), out var parsed)
            ? parsed
            : DateTimeOffset.MinValue;

    private static string RequireNodeString(JsonNode? node, string propertyName) =>
        NodeString(node, propertyName)
        ?? throw new InvalidOperationException($"Microsoft Graph response has no '{propertyName}'.");

    private static string? NodeString(JsonNode? node, string propertyName) =>
        node?[propertyName]?.GetValue<string>();

    private static DateTimeOffset NormalizeUtc(string value, string parameterName)
    {
        if (!DateTimeOffset.TryParse(value, out var parsed))
        {
            throw new ArgumentException($"{parameterName} must be an ISO 8601 timestamp.", parameterName);
        }

        return parsed.ToUniversalTime();
    }
}
