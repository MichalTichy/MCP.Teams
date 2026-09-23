using System.ComponentModel;
using ModelContextProtocol.Server;
using TeamsMcp.Services;

namespace TeamsMcp.Tools;

[McpServerToolType]
public sealed class MeetingTools(GraphClient graph)
{
    [McpServerTool(Name = "create_teams_meeting"), Description("Create a calendar-backed Microsoft Teams meeting and return the event including its join URL. This causes an external side effect and must only be called from an approved execution plan.")]
    public async Task<string> CreateTeamsMeeting(
        [Description("Meeting subject.")] string subject,
        [Description("ISO 8601 local date/time, for example 2026-09-15T14:00:00.")] string start,
        [Description("ISO 8601 local date/time.")] string end,
        [Description("Windows or IANA time-zone name accepted by Microsoft Graph, for example Europe/Prague.")] string timeZone,
        [Description("Required attendee email addresses.")] string[]? attendees = null,
        [Description("Optional HTML agenda or description.")] string? body = null,
        [Description("Optional reminder lead time in whole minutes. When supplied, the calendar reminder is enabled.")] int? reminderMinutesBeforeStart = null)
    {
        GraphClient.Require(subject, nameof(subject));
        GraphClient.Require(start, nameof(start));
        GraphClient.Require(end, nameof(end));
        GraphClient.Require(timeZone, nameof(timeZone));
        ValidateReminder(reminderMinutesBeforeStart);
        var payload = new Dictionary<string, object?>
        {
            ["subject"] = subject,
            ["body"] = new { contentType = "HTML", content = body ?? string.Empty },
            ["start"] = new { dateTime = start, timeZone },
            ["end"] = new { dateTime = end, timeZone },
            ["attendees"] = (attendees ?? []).Where(value => !string.IsNullOrWhiteSpace(value)).Select(address =>
                new { emailAddress = new { address = address.Trim() }, type = "required" }).ToArray(),
            ["isOnlineMeeting"] = true,
            ["onlineMeetingProvider"] = "teamsForBusiness",
            ["transactionId"] = Guid.NewGuid().ToString()
        };
        if (reminderMinutesBeforeStart is not null)
        {
            payload["isReminderOn"] = true;
            payload["reminderMinutesBeforeStart"] = reminderMinutesBeforeStart.Value;
        }

        return GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Post, "me/events", payload));
    }

    [McpServerTool(Name = "get_teams_meeting"), Description("Get a calendar-backed Teams meeting by event ID, including its online meeting details.")]
    public async Task<string> GetTeamsMeeting([Description("Calendar event ID returned by create_teams_meeting.")] string eventId) =>
        GraphClient.Print(await graph.DelegatedAsync(HttpMethod.Get, $"me/events/{GraphClient.Id(eventId)}"));

    [McpServerTool(Name = "update_teams_meeting"), Description("Update the subject, times, attendees, or agenda of a calendar-backed Teams meeting. Omitted fields remain unchanged. This causes an external side effect.")]
    public async Task<string> UpdateTeamsMeeting(
        [Description("Calendar event ID.")] string eventId,
        string? subject = null,
        string? start = null,
        string? end = null,
        string? timeZone = null,
        string[]? attendees = null,
        string? body = null,
        [Description("Optional reminder lead time in whole minutes. When supplied, the calendar reminder is enabled.")] int? reminderMinutesBeforeStart = null)
    {
        ValidateReminder(reminderMinutesBeforeStart);
        var payload = new Dictionary<string, object?>();
        if (!string.IsNullOrWhiteSpace(subject)) payload["subject"] = subject;
        if (body is not null) payload["body"] = new { contentType = "HTML", content = body };
        if (start is not null || end is not null)
        {
            GraphClient.Require(timeZone ?? string.Empty, nameof(timeZone));
            if (start is not null) payload["start"] = new { dateTime = start, timeZone };
            if (end is not null) payload["end"] = new { dateTime = end, timeZone };
        }
        if (attendees is not null)
            payload["attendees"] = attendees.Where(value => !string.IsNullOrWhiteSpace(value)).Select(address =>
                new { emailAddress = new { address = address.Trim() }, type = "required" }).ToArray();
        if (reminderMinutesBeforeStart is not null)
        {
            payload["isReminderOn"] = true;
            payload["reminderMinutesBeforeStart"] = reminderMinutesBeforeStart.Value;
        }
        if (payload.Count == 0) throw new ArgumentException("At least one meeting field must be supplied.");
        return GraphClient.Print(await graph.DelegatedAsync(new HttpMethod("PATCH"),
            $"me/events/{GraphClient.Id(eventId)}", payload));
    }

    [McpServerTool(Name = "cancel_teams_meeting"), Description("Cancel a calendar-backed Teams meeting and notify attendees. This causes an external side effect.")]
    public async Task<string> CancelTeamsMeeting(
        [Description("Calendar event ID.")] string eventId,
        [Description("Cancellation note sent to attendees.")] string comment = "") => GraphClient.Print(
        await graph.DelegatedAsync(HttpMethod.Post, $"me/events/{GraphClient.Id(eventId)}/cancel", new { comment }));

    private static void ValidateReminder(int? reminderMinutesBeforeStart)
    {
        if (reminderMinutesBeforeStart is < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(reminderMinutesBeforeStart),
                "Reminder lead time cannot be negative.");
        }
    }
}
