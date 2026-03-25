---
name: m365-copilot
description: "Query Microsoft 365 Copilot Chat API to search enterprise M365 data (emails, SharePoint, Teams, calendar) via natural language. Use when the user asks to search their M365 data, query Copilot about work documents, emails, meetings, or any Microsoft 365 content. Triggers on: M365 search, Copilot query, work emails, SharePoint docs, Teams messages, enterprise search."
---

# M365 Copilot Chat API Skill

Query Microsoft 365 Copilot programmatically through the Graph beta Chat API.
Copilot searches the user's M365 data (mail, SharePoint, Teams, calendar) respecting their permissions.

## Prerequisites

1. **Environment variables** must be set:
   - `M365_CLIENT_ID` — App Registration client ID from Entra ID
   - `M365_TENANT_ID` — Azure AD tenant ID
   - `M365_TOKEN_CACHE` — (optional) token cache path, defaults to `~/.m365_token_cache.json`

2. **Python dependencies**: `msal` and `requests`
   ```bash
   pip3 install msal requests
   ```

3. **First-time auth**: Run the auth command once to get a delegated token via device code flow. The user must visit a URL and enter a code. Token is cached and auto-refreshes.

## Usage

All commands via `scripts/m365_copilot.py`:

### Authenticate (one-time setup)
```bash
python3 scripts/m365_copilot.py auth
```
Displays a device code URL + code. User completes auth in browser. Token cached for future calls.

### Ask a question (new conversation)
```bash
python3 scripts/m365_copilot.py ask "What were the key decisions from last week's project meeting?"
```

### Continue a conversation
```bash
python3 scripts/m365_copilot.py chat <conversationId> "Can you elaborate on the budget discussion?"
```

### Options
- `--no-web` — Disable web search grounding (enterprise data only)
- `--context "text"` — Provide additional context for the query

### Output format
JSON with `answer`, `citations[]` (title, url, snippet), and `conversationId`.

## Workflow

1. Check if auth token exists (run `auth` if not)
2. Use `ask` to create a new conversation and send the query
3. Parse the JSON response — `answer` contains Copilot's response, `citations` has source references
4. For follow-up questions, use `chat` with the returned `conversationId`
5. Present answer and citations to the user

## Important Notes

- **Delegated only**: Runs in the context of the authenticated user. Copilot respects their M365 permissions.
- **Preview API**: Endpoints under `/beta` may change.
- **Text-only**: Copilot returns text responses only (no file creation, no email sending).
- **Web grounding**: Enabled by default. Use `--no-web` for enterprise-only results.
- **Rate limits**: Be mindful of Graph API throttling. Avoid rapid-fire queries.

## API Reference

For endpoint details, request/response schemas, and permissions: see `references/api-reference.md`.
