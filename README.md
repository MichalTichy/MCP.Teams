# MCP.Teams

A small, single-user Streamable HTTP MCP server for Microsoft Teams. It gives Samwise controlled access to chats, channels, calendar-backed Teams meetings, and recent call history through Microsoft Graph.

The service deliberately separates authentication by capability:

- Chat, channel, and meeting tools use a delegated Microsoft account connected with device-code sign-in.
- Call-history and transcript tools use application credentials because Microsoft Graph exposes these APIs through application permissions.
- Codex Gateway remains the authorization boundary. Only grant mutation tools to Samwise's execution profile.

## Microsoft Entra application

Create one app registration in the tenant that owns the Teams account.

1. Copy its tenant ID and application (client) ID into the environment file.
2. Under **Authentication**, enable **Allow public client flows**. This is required for the private device-code login used by the admin page.
3. Add delegated Microsoft Graph permissions:
   `User.Read`, `Chat.ReadWrite`, `ChatMessage.Send`, `Team.ReadBasic.All`,
   `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, and
   `Calendars.ReadWrite`.
4. Add the `CallRecords.Read.All`, `OnlineMeetingTranscript.Read.All`, and
   `CallTranscripts.Read.All` **application** permissions and grant admin consent.
5. Create a client secret if call history or transcript intake is needed. The delegated Teams tools do not use it.

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
  `teams_reply_to_channel_message`, `create_teams_meeting`,
  `update_teams_meeting`, and `cancel_teams_meeting` as needed.

Meeting creation is a calendar operation with `isOnlineMeeting=true`; this produces both a calendar event and a Teams join URL. It is intentionally different from a standalone Graph online meeting, which does not create a calendar event.

## Tool surface

- Chats: list/get chats, list/get messages, list members, send a message.
- Channels: list teams/channels/messages, send a root message, reply in a thread.
- Meetings: create, get, update, and cancel calendar-backed Teams meetings.
- Calls: list/get call records, participants, and sessions.
- Transcripts: list recent transcript metadata without downloading content, and
  download one selected meeting or ad-hoc call transcript on demand.

Message edits/deletes and team membership administration are intentionally omitted from this first version.
