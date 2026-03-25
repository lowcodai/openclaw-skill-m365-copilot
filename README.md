# 🤖 OpenClaw M365 Copilot Skill

An [OpenClaw](https://github.com/openclaw/openclaw) skill that connects your AI assistant to **Microsoft 365 Copilot** via the [Graph Chat API (beta)](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview).

Query your M365 data — emails, SharePoint, Teams, calendar, OneDrive — through natural language, all from your OpenClaw assistant.

## ✨ What it does

- **Search enterprise data** — Ask questions about your emails, documents, meetings, Teams messages
- **Multi-turn conversations** — Follow up on previous questions within the same conversation
- **Enterprise-grade security** — Uses delegated auth; Copilot respects your M365 permissions
- **Web + Enterprise grounding** — Toggle web search on/off per query

## 📋 Prerequisites

- **Microsoft 365 Copilot license** (E3/E5 + Copilot add-on)
- **OpenClaw** installed and running
- **Python 3.10+** with `msal` and `requests`
- **Entra ID App Registration** with required permissions

## 🚀 Quick Start

### 1. Install dependencies

```bash
pip3 install msal requests
```

### 2. Create App Registration in Entra ID

1. Go to **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name: `OpenClaw-Copilot` (or your choice)
3. Supported account types: **Single tenant**
4. Under **Authentication**:
   - Add platform: **Mobile and desktop applications**
   - Add redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Enable **"Allow public client flows"** → Yes
5. Under **API permissions**, add these **Delegated** Microsoft Graph permissions:
   - `Sites.Read.All`
   - `Mail.Read`
   - `People.Read.All`
   - `OnlineMeetingTranscript.Read.All`
   - `Chat.Read`
   - `ChannelMessage.Read.All`
   - `ExternalItem.Read.All`
6. **Grant admin consent** for your organization

### 3. Configure environment

```bash
cp .env.example .env
# Edit .env with your Client ID and Tenant ID
chmod 600 .env
```

### 4. Install the skill in OpenClaw

Add the skill directory to your OpenClaw config:

```json
{
  "skills": {
    "load": {
      "extraDirs": ["/path/to/your/skills"]
    }
  }
}
```

### 5. Authenticate

```bash
# Load env vars and run auth
export $(cat .env | xargs)
python3 scripts/m365_copilot.py auth
```

Follow the device code flow instructions — visit the URL and enter the code. Token is cached and auto-refreshes.

### 6. Test

```bash
python3 scripts/m365_copilot.py ask "What meetings do I have this week?"
```

## 📖 Usage

### CLI

```bash
# New question
python3 scripts/m365_copilot.py ask "Find documents about project X"

# Continue conversation
python3 scripts/m365_copilot.py chat <conversationId> "Tell me more about the budget"

# Enterprise data only (no web search)
python3 scripts/m365_copilot.py ask "Latest emails from John" --no-web

# Add extra context
python3 scripts/m365_copilot.py ask "Summarize this project" --context "Focus on Q1 deliverables"

# Custom timezone
python3 scripts/m365_copilot.py ask "My meetings tomorrow" --tz "US/Eastern"
```

### From OpenClaw

Once installed, just ask your OpenClaw assistant naturally:

> "What are my latest emails?"
> "Find SharePoint docs about the Q1 budget"
> "What meetings do I have tomorrow?"

OpenClaw will automatically use the skill to query M365 Copilot.

## 🔧 Configuration

| Env Variable | Required | Default | Description |
|---|---|---|---|
| `M365_CLIENT_ID` | ✅ | — | App Registration client ID |
| `M365_TENANT_ID` | ✅ | — | Entra ID tenant ID |
| `M365_CLIENT_SECRET` | ❌ | — | Client secret (enables confidential client) |
| `M365_TOKEN_CACHE` | ❌ | `~/.m365_token_cache.json` | Token cache file path |
| `M365_TIMEZONE` | ❌ | `UTC` | Default timezone for queries |
| `M365_REDIRECT_URI` | ❌ | `http://localhost:19365/auth/callback` | Redirect URI for auth code flow |

## ⚠️ Known Limitations

These are **Microsoft Copilot Chat API** limitations (not this skill):

- **Text-only responses** — No image generation, code interpreter, or file creation
- **No actions** — Can't send emails, create events, or modify documents
- **Preview API** — Endpoints under `/beta` may change
- **Copilot license required** — Each user needs a M365 Copilot add-on
- **Web search toggle is per-turn** — Must be set on each message

## 📁 Structure

```
m365-copilot/
├── SKILL.md              # OpenClaw skill definition
├── scripts/
│   └── m365_copilot.py   # CLI client (auth, ask, chat)
├── references/
│   └── api-reference.md  # API documentation
├── .env.example          # Environment template
├── LICENSE               # MIT
└── README.md             # This file
```

## 🤝 Contributing

1. Fork the repo
2. Create a feature branch
3. Test with your own M365 tenant
4. Submit a PR

## 📄 License

MIT — See [LICENSE](LICENSE)

## 🔗 Links

- [OpenClaw](https://github.com/openclaw/openclaw)
- [ClawHub Skills Marketplace](https://clawhub.com)
- [M365 Copilot Chat API Docs](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview)
