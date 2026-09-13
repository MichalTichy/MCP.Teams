using System.ComponentModel;
using ModelContextProtocol.Server;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class CallHistoryTools(GraphClient graph)
{
    [McpServerTool(Name = "teams_call_history_list_records"), Description("List recent Teams peer-to-peer, group-call, and meeting call records. Microsoft Graph retains call records for a limited period.")]
    public async Task<string> ListRecords(
        [Description("Optional inclusive UTC ISO 8601 start timestamp.")] string? fromDateTime = null,
        [Description("Optional exclusive UTC ISO 8601 end timestamp.")] string? toDateTime = null,
        [Description("Maximum records to return (1-60).")]
        int limit = 20)
    {
        var filters = new List<string>();
        if (!string.IsNullOrWhiteSpace(fromDateTime)) filters.Add($"startDateTime ge {NormalizeUtc(fromDateTime)}");
        if (!string.IsNullOrWhiteSpace(toDateTime)) filters.Add($"startDateTime lt {NormalizeUtc(toDateTime)}");
        var query = $"?$top={Math.Clamp(limit, 1, 60)}";
        if (filters.Count > 0) query += $"&$filter={Uri.EscapeDataString(string.Join(" and ", filters))}";
        return GraphClient.Print(await graph.ApplicationAsync(HttpMethod.Get, $"communications/callRecords{query}"));
    }

    [McpServerTool(Name = "teams_call_history_get_record"), Description("Get one Teams call record by ID.")]
    public async Task<string> GetRecord([Description("Call record ID.")] string callRecordId) =>
        GraphClient.Print(await graph.ApplicationAsync(HttpMethod.Get,
            $"communications/callRecords/{GraphClient.Id(callRecordId)}"));

    [McpServerTool(Name = "teams_call_history_list_participants"), Description("List participants for a Teams call record.")]
    public async Task<string> ListParticipants([Description("Call record ID.")] string callRecordId) =>
        GraphClient.Print(await graph.ApplicationAsync(HttpMethod.Get,
            $"communications/callRecords/{GraphClient.Id(callRecordId)}/participants_v2"));

    [McpServerTool(Name = "teams_call_history_list_sessions"), Description("List media sessions for a Teams call record.")]
    public async Task<string> ListSessions(
        [Description("Call record ID.")] string callRecordId,
        [Description("Include media segments in each session.")] bool includeSegments = false)
    {
        var suffix = includeSegments ? "?$expand=segments" : string.Empty;
        return GraphClient.Print(await graph.ApplicationAsync(HttpMethod.Get,
            $"communications/callRecords/{GraphClient.Id(callRecordId)}/sessions{suffix}"));
    }

    private static string NormalizeUtc(string value)
    {
        if (!DateTimeOffset.TryParse(value, out var parsed)) throw new ArgumentException($"Invalid UTC date/time: {value}");
        return parsed.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
    }
}
