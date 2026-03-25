# M365 Copilot Chat API Reference

## Endpoints (Beta)

Base: `https://graph.microsoft.com/beta`

### Create Conversation
```
POST /copilot/conversations
Body: {}
Response: 201 Created → { "id": "<conversationId>" }
```

### Send Message (Synchronous)
```
POST /copilot/conversations/{conversationId}/chat
Headers:
  Authorization: Bearer {token}
  Content-Type: application/json
Body:
{
  "message": {
    "content": "Your question here"
  },
  "additionalContext": [           // optional
    {
      "content": "extra context",
      "contentType": "reference"   // or "inlineReference"
    }
  ],
  "contextualResources": {         // optional
    "webContext": {
      "isWebEnabled": false        // default true, set false to disable per-turn
    },
    "files": [                     // optional — OneDrive/SharePoint file URIs
      { "uri": "https://contoso.sharepoint.com/..." }
    ]
  }
}
Response: 200 OK → copilotConversationResponse
```

### Send Message (Streamed)
```
POST /copilot/conversations/{conversationId}/chatOverStream
Same body as synchronous. Returns chunked response.
```

## Required Permissions (ALL delegated, work/school only)

- Sites.Read.All
- Mail.Read
- People.Read.All
- OnlineMeetingTranscript.Read.All
- Chat.Read
- ChannelMessage.Read.All
- ExternalItem.Read.All

All 7 are required simultaneously. Application-only NOT supported.

## Response Format

```json
{
  "id": "message-id",
  "content": "Copilot's answer text",
  "citations": [
    {
      "title": "Document title",
      "url": "https://...",
      "snippet": "relevant excerpt"
    }
  ]
}
```

## Known Limitations

- No action/content generation (can't create files, send emails, schedule meetings)
- Text-only responses (no images, code interpreter, graphic art)
- No long-running tasks (gateway timeout risk)
- Web search grounding toggle is per-turn only
- Subject to semantic index limitations
- Preview API: may change without notice

## OAuth2 Delegated Flow

Uses MSAL (Microsoft Authentication Library):
1. Register app in Entra ID (Azure AD)
2. Configure redirect URI (http://localhost for device/auth code flow)
3. Request token with all 7 scopes
4. Token includes user context → Copilot respects user's M365 permissions

### Scopes String
```
Sites.Read.All Mail.Read People.Read.All OnlineMeetingTranscript.Read.All Chat.Read ChannelMessage.Read.All ExternalItem.Read.All
```

### Token Cache
MSAL supports serialized token cache. Store in secure location.
Refresh tokens auto-renew access tokens silently.
