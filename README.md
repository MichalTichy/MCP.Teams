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
   `Calendars.ReadWrite`, `Files.ReadWrite`, and `OnlineMeetingTranscript.Read.All`.
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
  or messages with new file uploads and existing OneDrive/SharePoint references.
- Channels: list teams/channels/messages, send a root message, reply in a thread.
- Meetings: create, get, update, and cancel calendar-backed Teams meetings.
- Calls: list/get call records, participants, and sessions.
- Transcripts: list recent transcript metadata for organized meetings, invited
  calendar meetings, and ad-hoc calls without downloading content, then download
  one selected transcript on demand. Discovery returns successful sources together
  with per-source errors instead of dropping all results when one Graph API fails.

Message edits/deletes and team membership administration are intentionally omitted from this first version.

## Chat file attachments

To deliver a message or files privately to the connected account, pass `chatId: "self"`
to the existing chat tools. This resolves to Teams' `48:notes` self-chat. No chat URL,
participant search, or chat creation is needed, even when it is absent from the recent
chat list. `teams_list_chats` also returns this target in `selfChat`; `teams_get_chat`
and `teams_list_chat_members` resolve it from the connected profile instead of relying
on Graph's ordinary self-chat metadata or membership endpoints. File uploads stay in
the connected account's OneDrive with no sharing invitations. Other chats continue to
require verified membership and sharing with every recipient.

`teams_send_chat_message_with_attachments` is a separate mutation tool; the existing
`teams_send_chat_message` contract is unchanged. It accepts `chatId`, a plain-text
`message`, and 1–10 files in total using either or both optional arrays:

- `files`: new files to upload, share, and send in this single call. Each has
  `name`, optional `contentType`, and exactly one of `agentFileUrl` or
  `base64Content`. The default download limit is 25 MiB per file, 50 MiB combined;
  inline base64 is limited to 1 MiB decoded. Sources are read and validated before
  making changes to OneDrive.
- `attachments`: existing, already shared references with `id` (SharePoint GUID),
  `contentUrl` (HTTPS `webDavUrl`), and `name`. This preserves the previous contract;
  reference-only calls make no upload or sharing requests.

For example, to send a generated workbook through Codex Gateway:

```json
{
  "chatId": "<verified existing chat ID>",
  "message": "Here are the working reports for review.",
  "files": [
    {
      "name": "report.xlsx",
      "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "agentFileUrl": "./artifacts/report.xlsx"
    }
  ]
}
```

Gateway snapshots `./artifacts/...` into a temporary URL before forwarding the
authorized tool call. Direct MCP clients can supply an existing Gateway artifact
URL or inline base64 instead. `TEAMS_MCP_GATEWAY_FILE_BASE_URL` must match the
origin in Gateway's `McpGateway:RunnerBaseUrl` (Compose: `http://gateway:8080`).
Only opaque `/_internal/mcp/artifacts/...` URLs on that exact origin are accepted;
downloads have no Graph credentials and do not follow redirects.

The tool resolves the connected account and all chat members, including later
pages. It uploads new files into the connected account's `Microsoft Teams Chat
Files` folder, creating the folder when needed. Unique filenames preserve existing
files. Every other chat member gets authenticated read access without invitation
emails or public sharing links. A self-chat needs no additional sharing. Unsupported
or unresolved member identities stop delivery. All sharing receipts must confirm
the intended recipients before one message is posted with all file references.

Failures return `isError: true`, prepared `attachments`, and `uploadedFiles` with
their drive item IDs, names, available URLs, and confirmed-sharing status. Before
the message request, `messageOutcome` is `notSent`; uploaded files may remain and
are not automatically deleted. An upload with a lost receipt can also leave a
file without a recorded item ID. A missing message receipt returns `unknown`;
inspect the chat before retrying. A failure after receiving the message ID reports
`sent`. Mutations are never automatically retried. Successful sends publish a
`ChatFilesMessageSentNotification`. Existing-reference callers remain responsible
for those files' recipient permissions.

To enable the feature on an existing installation:

1. Add delegated `Files.ReadWrite` to the Entra registration and grant the required
   consent. Deploy the updated Teams MCP image, then reconnect the same Microsoft
   account through `/admin` so the new permission enters the token cache. Existing
   custom `DelegatedScopes` settings must include `Files.ReadWrite` for uploads.
2. Deploy the updated Gateway for `./artifacts/...` output-file forwarding and
   HTTPS binary input attachments. Deploy its runner with Python/openpyxl and the
   document extractor when those are not already available. Check the file-origin
   setting above. A separate OneDrive MCP server is not needed.
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

Tests use fake HTTP transports and token providers. They cover upload bytes,
Gateway downloads, recipient resolution and pagination, sharing failures, self-chat,
folder creation, source validation, partial delivery, existing file references,
HTML escaping, the actual MCP schema, and the existing text-only contract. They do
not contact Microsoft Graph or send real messages.
