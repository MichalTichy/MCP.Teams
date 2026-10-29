# MCP.Teams

A small, single-user Streamable HTTP MCP server for Microsoft Teams. It gives Samwise controlled access to chats, channels, calendar-backed Teams meetings, and recent call history through Microsoft Graph.

The service deliberately separates authentication by capability:

- Chat, channel, and meeting tools use a delegated Microsoft account connected with device-code sign-in.
- Call-history and transcript discovery use application credentials. Scheduled-meeting transcript downloads first use the connected account's delegated access, which also supports meetings the user joined but did not organize, and then fall back to application access. Ad-hoc call transcripts use application access.
- Codex Gateway remains the authorization boundary. Only grant mutation tools to Samwise's execution profile.

## Microsoft Entra application

Create one app registration in the tenant that owns the Teams account.

1. Copy its tenant ID and application (client) ID into the environment file.
2. Under **Authentication**, enable **Allow public client flows**. This is required for the private device-code login used by the admin page.
3. Add delegated Microsoft Graph permissions:
   `User.Read`, `Chat.ReadWrite`, `ChatMessage.Send`, `Team.ReadBasic.All`,
   `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`,
   `Calendars.ReadWrite`, and `OnlineMeetingTranscript.Read.All`.
4. Add the `CallRecords.Read.All`, `OnlineMeetingTranscript.Read.All`, and
   `CallTranscripts.Read.All` **application** permissions and grant admin consent.
5. Create a client secret if call history, transcript discovery, or ad-hoc call transcript intake is needed.

The exact permissions available can depend on the Microsoft 365 tenant and its policies. Admin consent may also be required for delegated permissions.

## Run

Copy `.env.example` to `.env`, fill the settings, then:

```powershell
docker compose up --build -d
```

Open `http://localhost:5052/admin`, enter `TEAMS_MCP_ADMIN_PASSWORD`, and choose **Connect Microsoft account**. Open the Microsoft verification URL, enter the displayed code, and finish sign-in. The persistent volume retains the MSAL token cache; treat that volume as a secret.

The container joins `codex-gateway-backend` and still has normal outbound internet access to Microsoft login and Graph endpoints. Its host port exists only for the private admin UI and can be bound to a specific LAN address instead of `0.0.0.0`.

## Gateway registration

Add an HTTP MCP server in Codex Gateway:

| Field | Value |
| --- | --- |
| Name | `Microsoft Teams` |
| URL | `http://teams-mcp:8080` |
| Execution mode | `Gateway` |
| Headers | none |

Discover the tools, then grant them by Samwise stage:

- Intake, planning, and execution: grant the read-only chat, channel, meeting,
  transcript, and call-history tools each stage actually needs. Transcript input
  requires `teams_get_transcript` on the intake profile.
- Execution additionally: `teams_send_chat_message`, `teams_send_channel_message`,
  `teams_send_chat_message_with_attachments`, `teams_reply_to_channel_message`, `create_teams_meeting`,
  `update_teams_meeting`, and `cancel_teams_meeting` as needed.

Meeting creation is a calendar operation with `isOnlineMeeting=true`; this produces both a calendar event and a Teams join URL. It accepts an optional `reminderMinutesBeforeStart` value and enables the event reminder when supplied. It is intentionally different from a standalone Graph online meeting, which does not create a calendar event.

## Tool surface

- Chats: list/get chats, list/get messages, list members, send plain-text messages
  or messages with references to existing OneDrive/SharePoint files.
- Channels: list teams/channels/messages, send a root message, reply in a thread.
- Meetings: create, get, update, and cancel calendar-backed Teams meetings.
- Calls: list/get call records, participants, and sessions.
- Transcripts: list recent transcript metadata for organized meetings, invited
  calendar meetings, and ad-hoc calls without downloading content, then download
  one selected transcript on demand. Discovery returns successful sources together
  with per-source errors instead of dropping all results when one Graph API fails.

Message edits/deletes and team membership administration are intentionally omitted from this first version.

## Chat file attachments

`teams_send_chat_message_with_attachments` is a separate mutation tool; the existing
`teams_send_chat_message` contract is unchanged. It accepts `chatId`, a plain-text
`message`, and 1–10 references to existing Microsoft 365 files. It only posts a
Teams message: it does not upload, download, modify, or share files. Upload and
recipient permissions belong to a separate OneDrive/SharePoint tool.

For each attachment, pass the storage tool's verified metadata:

- `id`: the SharePoint file GUID from `sharepointIds.listItemUniqueId` or the
  GUID in the driveItem `eTag`. Do not use the opaque driveItem ID.
- `contentUrl`: the file's HTTPS `webDavUrl`, selected explicitly when reading
  driveItem metadata. This is a Microsoft 365 file reference, not a Gateway
  artifact URL, Graph API download endpoint, base64 payload, or local path.
- `name`: the display filename including its extension.

For example:

```json
{
  "chatId": "<verified existing chat ID>",
  "message": "Here are the working reports for review.",
  "attachments": [
    {
      "id": "153fa47d-18c9-4179-be08-9879815a9f90",
      "contentUrl": "https://tenant.sharepoint.com/personal/owner/Documents/Reports/report.xlsx",
      "name": "report.xlsx"
    }
  ]
}
```

All intended recipients must already have access. The separate storage tool must
verify that prerequisite; this Teams tool neither grants nor checks file access.
Its only Microsoft Graph request is `POST /chats/{chatId}/messages`, with
`contentType: reference` attachments. Text is HTML-escaped, references are validated
before sending, and GUIDs must be unique within a message. It uses the existing
delegated `ChatMessage.Send` permission and requires no file permissions or Gateway
file-transfer configuration.

Failures before sending return `isError: true` without external changes. A send
attempt that ends without a valid message receipt returns `isError: true`, the
file references, and `messageOutcome: unknown`. Inspect the chat before retrying;
the tool never automatically retries a mutation. A failure after receiving the
message ID is reported as `messageOutcome: sent` so an already sent message is not
duplicated. Successful sends publish a `ChatFilesMessageSentNotification`.

To enable the feature on an existing installation:

1. Make a separate storage tool available to upload files, establish recipient
   access, and return the reference metadata above. That capability is not
   implemented by this Teams server.
2. Build/redeploy the Teams MCP image. No new Graph scopes are required by this
   change and the existing delegated scope set is unchanged.
3. Rediscover this server's tools in Gateway and grant invocation of
   `teams_send_chat_message_with_attachments` only to the execution profile.
   Samwise planning can inspect it through the Execution catalogue/schema lookup;
   it must not invoke it during planning.
4. Replan the pending Samwise item after the new tool is available.

Reference APIs: [Teams file references](https://learn.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0#example-4-send-a-message-with-file-attachment-in-it)
and [chat sending permissions](https://learn.microsoft.com/en-us/graph/api/chat-post-messages?view=graph-rest-1.0).

## Development checks

```powershell
dotnet build TeamsMcp.csproj --configuration Release
dotnet test tests/TeamsMcp.Tests.csproj --configuration Release
```

Tests use fake HTTP transports and token providers. They verify that existing
file references produce exactly one Teams send and no storage/sharing operations,
validate references and HTML escaping, handle uncertain sends, inspect the actual
MCP schema, and preserve the existing text-only contract. They do not contact
Microsoft Graph or send real messages.
