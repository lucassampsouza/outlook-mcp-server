# Outlook Calendar MCP Server — Setup Guide

## Prerequisites

- Python 3.10+
- A Microsoft 365 / Azure AD account with admin rights (or an admin who can consent for you)

---

## 1. Create an Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**.
2. Name it (e.g. `outlook-mcp-server`). Leave redirect URI blank. Click **Register**.
3. Note down:
   - **Application (client) ID** → `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → `AZURE_TENANT_ID`

---

## 2. Add API Permissions

1. In your App Registration, go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
2. Add the following permissions:
   | Permission | Purpose |
   |---|---|
   | `Calendars.Read` | Read calendar events |
   | `Calendars.ReadWrite` | Create / update / delete events |
   | `MailboxSettings.Read` | Read timezone settings (optional) |
3. Click **Grant admin consent** (requires a Global Admin or Privileged Role Admin).

> **Note:** Application permissions allow the server to act on behalf of *any* user in the tenant.
> If you only want access to specific mailboxes, use **delegated permissions** and implement
> the Authorization Code or Device Code OAuth flow instead of client credentials.

---

## 3. Create a Client Secret

1. In your App Registration, go to **Certificates & secrets** → **New client secret**.
2. Set an expiry, click **Add**.
3. **Copy the secret _value_ immediately** — it won't be shown again.
   This is your `AZURE_CLIENT_SECRET`.

---

## 4. Configure Environment Variables

```bash
cp .env.example .env
# Edit .env and fill in AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
```

---

## 5. Install Dependencies

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## 6. Run the Server

### Stdio mode (default — for use with Claude Desktop or other MCP clients)

```bash
python server.py
```

### HTTP/SSE mode

```bash
fastmcp run server.py --transport sse --port 8000
```

---

## 7. Connect to Claude Desktop

Add this block to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-calendar": {
      "command": "python",
      "args": ["/absolute/path/to/outlook-mcp-server/server.py"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "..."
      }
    }
  }
}
```

---

## Multi-account Setup

If you have **two separate Microsoft accounts in different Azure AD tenants**, you can
configure them as named profiles. The server loads all accounts on startup; each tool
accepts an optional `account` parameter to select which credentials to use.

### Step 1 — Create a second App Registration

Repeat steps 1–3 above in the **second** Azure tenant. You will get a second set of
`TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET`.

### Step 2 — Add named account variables to `.env`

Named accounts follow the pattern `ACCOUNT_{NAME}_TENANT_ID` / `_CLIENT_ID` / `_CLIENT_SECRET`.
The `{NAME}` part (case-insensitive) becomes the value you pass to tools.

```dotenv
# Primary / default account (unchanged)
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=primary-secret

# Second account named "personal"
ACCOUNT_PERSONAL_TENANT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_PERSONAL_CLIENT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_PERSONAL_CLIENT_SECRET=personal-secret
```

You can add as many named accounts as you need.

### Step 3 — Use the `account` parameter in tools

Every tool now accepts an optional `account` parameter (default `"default"`):

```
list_accounts()
  → {"accounts": ["default", "personal"], ...}

list_calendars(user_email="alice@work.com")
  → uses default account

list_calendars(user_email="bob@personal.com", account="personal")
  → uses personal account credentials

get_calendar_events(user_email="bob@personal.com", account="personal")
create_event(user_email="bob@personal.com", subject="Dinner", ..., account="personal")
```

Tokens for each account are cached independently, so switching between accounts
within a session is efficient.

### Multi-account Claude Desktop config

You can pass the extra env vars directly in the Claude Desktop config:

```json
{
  "mcpServers": {
    "outlook-calendar": {
      "command": "python",
      "args": ["/absolute/path/to/outlook-mcp-server/server.py"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "ACCOUNT_PERSONAL_TENANT_ID": "...",
        "ACCOUNT_PERSONAL_CLIENT_ID": "...",
        "ACCOUNT_PERSONAL_CLIENT_SECRET": "..."
      }
    }
  }
}
```

---

## Available Tools

| Tool | Description |
|---|---|
| `list_accounts` | List all configured account names |
| `list_calendars` | List all calendars for a user |
| `get_calendar_events` | Fetch events in a time window |
| `get_event` | Get full details for a specific event |
| `create_event` | Create a new event (optionally with Teams link) |
| `update_event` | Patch fields on an existing event |
| `delete_event` | Delete an event |
| `get_free_busy` | Get free/busy schedule for one or more users |
| `search_events` | Search events by keyword |

All tools that read or write user data require the user's **email address** as the first
argument. All tools accept an optional `account` parameter (default `"default"`) to
select which configured Azure AD account to authenticate with.
