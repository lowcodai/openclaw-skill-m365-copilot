#!/usr/bin/env python3
"""
M365 Copilot Chat API client for OpenClaw.
Supports both public client (device code flow) and confidential client (auth code flow).

Usage:
  python3 m365_copilot.py auth                      # Authenticate (one-time)
  python3 m365_copilot.py ask "query"                # New conversation + question
  python3 m365_copilot.py chat <convId> "query"      # Continue conversation
  python3 m365_copilot.py ask "query" --no-web       # Enterprise data only
  python3 m365_copilot.py ask "query" --context "..." # Extra grounding context
  python3 m365_copilot.py ask "query" --tz US/Eastern # Override timezone

Env vars (required):
  M365_CLIENT_ID     - App Registration client ID
  M365_TENANT_ID     - Azure AD / Entra ID tenant ID

Env vars (optional):
  M365_CLIENT_SECRET - Client secret (if set, uses confidential client + auth code flow)
  M365_TOKEN_CACHE   - Token cache file path (default: ~/.m365_token_cache.json)
  M365_TIMEZONE      - Default timezone (default: UTC)
  M365_REDIRECT_URI  - Redirect URI for auth code flow (default: http://localhost:19365/auth/callback)
"""

import argparse
import json
import os
import sys

try:
    import msal
except ImportError:
    print("ERROR: msal not installed. Run: pip3 install msal", file=sys.stderr)
    sys.exit(1)

try:
    import requests
except ImportError:
    print("ERROR: requests not installed. Run: pip3 install requests", file=sys.stderr)
    sys.exit(1)

GRAPH_BASE = "https://graph.microsoft.com/beta"
SCOPES = [
    "Sites.Read.All",
    "Mail.Read",
    "People.Read.All",
    "OnlineMeetingTranscript.Read.All",
    "Chat.Read",
    "ChannelMessage.Read.All",
    "ExternalItem.Read.All",
]

CACHE_PATH = os.environ.get("M365_TOKEN_CACHE", os.path.expanduser("~/.m365_token_cache.json"))
DEFAULT_TZ = os.environ.get("M365_TIMEZONE", "UTC")
REDIRECT_URI = os.environ.get("M365_REDIRECT_URI", "http://localhost:19365/auth/callback")


# ── MSAL helpers ──────────────────────────────────────────────────────────────

def get_msal_app():
    """Create MSAL client with serialized token cache."""
    client_id = os.environ.get("M365_CLIENT_ID")
    tenant_id = os.environ.get("M365_TENANT_ID")
    client_secret = os.environ.get("M365_CLIENT_SECRET")
    if not client_id or not tenant_id:
        print("ERROR: M365_CLIENT_ID and M365_TENANT_ID must be set", file=sys.stderr)
        sys.exit(1)

    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r") as f:
            cache.deserialize(f.read())

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    if client_secret:
        app = msal.ConfidentialClientApplication(
            client_id, authority=authority,
            client_credential=client_secret, token_cache=cache,
        )
    else:
        app = msal.PublicClientApplication(
            client_id, authority=authority, token_cache=cache,
        )
    return app, cache


def save_cache(cache):
    """Persist token cache to disk (chmod 600)."""
    if cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())
        os.chmod(CACHE_PATH, 0o600)


def get_token(app, cache, interactive=False):
    """Acquire token silently, or interactively if needed."""
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        if not interactive:
            print("ERROR: No cached token. Run 'm365_copilot.py auth' first.", file=sys.stderr)
            sys.exit(1)

        is_confidential = isinstance(app, msal.ConfidentialClientApplication)

        if is_confidential:
            result = _auth_code_flow(app)
        else:
            result = _device_code_flow(app)

    save_cache(cache)

    if not result or "access_token" not in result:
        err = result.get("error_description", json.dumps(result)) if result else "unknown error"
        print(f"ERROR: Auth failed: {err}", file=sys.stderr)
        sys.exit(1)

    return result["access_token"]


def _device_code_flow(app):
    """Public client: device code flow."""
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"ERROR: Device flow failed: {flow.get('error_description', 'unknown')}", file=sys.stderr)
        sys.exit(1)
    print(f"\n🔐 To authenticate, visit: {flow['verification_uri']}")
    print(f"   Enter code: {flow['user_code']}\n")
    return app.acquire_token_by_device_flow(flow)


def _auth_code_flow(app):
    """Confidential client: authorization code flow with manual redirect paste."""
    from urllib.parse import urlparse, parse_qs

    flow = app.initiate_auth_code_flow(scopes=SCOPES, redirect_uri=REDIRECT_URI)
    if "auth_uri" not in flow:
        print(f"ERROR: Auth code flow failed: {flow}", file=sys.stderr)
        sys.exit(1)
    print(f"\n🔐 Visit this URL to authenticate:")
    print(f"   {flow['auth_uri']}\n")
    print(f"After login, you'll be redirected to a page that won't load.")
    print(f"Copy the FULL URL from your browser's address bar and paste it here:\n")

    redirect_url = input("Paste URL here: ").strip()
    if not redirect_url:
        print("ERROR: No redirect URL provided", file=sys.stderr)
        sys.exit(1)

    parsed = urlparse(redirect_url)
    params = parse_qs(parsed.query)
    auth_response = {k: v[0] for k, v in params.items()}
    return app.acquire_token_by_auth_code_flow(flow, auth_response)


# ── Graph API helpers ─────────────────────────────────────────────────────────

def _headers(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def create_conversation(token):
    """Create a new Copilot conversation. Returns conversationId."""
    r = requests.post(f"{GRAPH_BASE}/copilot/conversations", headers=_headers(token), json={})
    if not r.ok:
        print(f"ERROR {r.status_code}: {r.text}", file=sys.stderr)
        r.raise_for_status()
    return r.json().get("id")


def send_message(token, conv_id, message, web_search=True, context=None, timezone=None):
    """Send a synchronous chat message to a Copilot conversation."""
    tz = timezone or DEFAULT_TZ
    body = {
        "message": {"text": message},
        "locationHint": {"timeZone": tz},
    }
    if not web_search:
        body.setdefault("contextualResources", {})["isWebSearchEnabled"] = False
    if context:
        body["additionalContext"] = [{"content": context, "contentType": "reference"}]

    r = requests.post(
        f"{GRAPH_BASE}/copilot/conversations/{conv_id}/chat",
        headers=_headers(token), json=body,
    )
    if not r.ok:
        print(f"ERROR {r.status_code}: {r.text}", file=sys.stderr)
        r.raise_for_status()
    return r.json()


def format_response(resp):
    """Extract answer text and citations from Copilot response."""
    messages = resp.get("messages", [])
    answer = ""
    citations = []

    for msg in messages:
        odata = msg.get("@odata.type", "")
        text = msg.get("text", "")
        # Take the last response message (skip user echo)
        if odata.endswith("ResponseMessage") and text:
            answer = text
            for a in msg.get("attributions", []):
                if a.get("attributionType") == "citation":
                    citations.append({
                        "title": a.get("providerDisplayName", ""),
                        "url": a.get("seeMoreWebUrl", ""),
                    })

    if not answer:
        answer = resp.get("text", resp.get("content", ""))

    output = {"answer": answer}
    if citations:
        output["citations"] = citations
    return output


# ── CLI commands ──────────────────────────────────────────────────────────────

def cmd_auth(args):
    app, cache = get_msal_app()
    get_token(app, cache, interactive=True)
    print(f"✅ Authenticated successfully. Token cached at {CACHE_PATH}")


def cmd_ask(args):
    app, cache = get_msal_app()
    token = get_token(app, cache)
    conv_id = create_conversation(token)
    resp = send_message(token, conv_id, args.query,
                        web_search=not args.no_web, context=args.context, timezone=args.tz)
    result = format_response(resp)
    result["conversationId"] = conv_id
    print(json.dumps(result, indent=2, ensure_ascii=False))


def cmd_chat(args):
    app, cache = get_msal_app()
    token = get_token(app, cache)
    resp = send_message(token, args.conv_id, args.query,
                        web_search=not args.no_web, context=args.context, timezone=args.tz)
    result = format_response(resp)
    result["conversationId"] = args.conv_id
    print(json.dumps(result, indent=2, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description="M365 Copilot Chat API client for OpenClaw")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("auth", help="Authenticate (device code or auth code flow)")

    for name, help_text in [("ask", "New conversation + question"), ("chat", "Continue conversation")]:
        p = sub.add_parser(name, help=help_text)
        if name == "chat":
            p.add_argument("conv_id", help="Conversation ID to continue")
        p.add_argument("query", help="Question or message to send")
        p.add_argument("--no-web", action="store_true", help="Disable web search grounding")
        p.add_argument("--context", help="Additional context text for grounding")
        p.add_argument("--tz", default=DEFAULT_TZ, help=f"Timezone (default: {DEFAULT_TZ})")

    args = parser.parse_args()
    {"auth": cmd_auth, "ask": cmd_ask, "chat": cmd_chat}[args.command](args)


if __name__ == "__main__":
    main()
