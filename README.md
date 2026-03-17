# Outlook Calendar MCP Server

An MCP server that exposes Microsoft Outlook calendar operations via the **Microsoft Graph API**. Works with Claude Desktop, Claude Code, and any MCP-compatible client.

## Features

- List, create, update, and delete calendar events
- Query free/busy availability for one or more users
- Search events by keyword
- Multi-account support (multiple Azure AD tenants/users simultaneously)
- Two authentication modes: **Application** (client secret) and **Delegated** (refresh token via Device Code flow)

## Available Tools

| Tool | Description |
|---|---|
| `list_accounts` | List all configured account names |
| `list_calendars` | List all calendars for a user |
| `get_calendar_events` | Fetch events within a time window |
| `get_event` | Get full details for a specific event |
| `create_event` | Create a new event (optionally with Teams link) |
| `update_event` | Patch fields on an existing event |
| `delete_event` | Delete an event |
| `get_free_busy` | Get free/busy schedule for one or more users |
| `search_events` | Search events by keyword |
| `start_device_code_auth` | Begin a Device Code auth flow to add a delegated account |
| `complete_device_code_auth` | Finish a pending Device Code auth and save the refresh token |
| `get_admin_consent_url` | Generate an admin consent URL for a tenant |

---

## 1. Azure App Registration

Before running the server, you need an Azure App Registration.

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Give it a name (e.g. `outlook-mcp-server`), leave redirect URI blank, and click **Register**
3. Note down:
   - **Application (client) ID** → `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → `AZURE_TENANT_ID`

### API Permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph**:

| Permission | Type | Purpose |
|---|---|---|
| `Calendars.Read` | Application or Delegated | Read calendar events |
| `Calendars.ReadWrite` | Application or Delegated | Create / update / delete events |
| `Calendars.Read.Shared` | Delegated | Access shared calendars |
| `User.Read` | Delegated | Read the signed-in user's profile |

- **Application permissions** require clicking **Grant admin consent** (needs a Global Admin)
- **Delegated permissions** require the user to authenticate via Device Code flow (no admin needed)

---

## 2. Authentication Modes

### Mode 1 — Application (Client Credentials)

Best for server-side or automated access to any mailbox in the tenant. Requires admin consent.

1. In the App Registration, go to **Certificates & secrets** → **New client secret**
2. Copy the secret **value** (not the ID) — this is `AZURE_CLIENT_SECRET`

Required env vars:
```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

### Mode 2 — Delegated (Device Code Flow)

Best for personal use or when admin consent is not available. No client secret needed.

Run the interactive setup script:
```bash
python setup.py
```

It will prompt for your Tenant ID and Client ID, open a browser-based Device Code flow, and save the refresh token to `.env` automatically.

Alternatively, authenticate directly through the MCP tools:
```
start_device_code_auth(account_name="default", tenant_id="...", client_id="...")
# → Returns a code and URL. Visit the URL, sign in, then:
complete_device_code_auth(account_name="default")
```

Required env vars (after setup):
```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_REFRESH_TOKEN=your-refresh-token
```

> The server automatically rotates the refresh token on every use and saves the new one to `.env`.

---

## 3. Installation & Running

### Option A — uvx (recommended, no local clone needed)

```bash
uvx --from git+https://github.com/lucassampsouza/outlook-mcp-server outlook-mcp-server
```

### Option B — Local clone

```bash
git clone https://github.com/lucassampsouza/outlook-mcp-server
cd outlook-mcp-server
cp .env.example .env
# Fill in .env with your credentials

python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python server.py
```

### HTTP/SSE mode

```bash
fastmcp run server.py --transport sse --port 8000
```

---

## 4. Connecting to Claude

### stdio (Claude Desktop / Claude Code)

Add to `claude_desktop_config.json` (usually at `~/Library/Application Support/Claude/` on macOS or `%APPDATA%\Claude\` on Windows):

**Using uvx — Application auth:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

**Using uvx — Delegated auth:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_REFRESH_TOKEN": "your-refresh-token"
      }
    }
  }
}
```

**Using local clone:**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["/absolute/path/to/outlook-mcp-server/server.py"],
      "env": {
        "AZURE_TENANT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_ID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### HTTP/SSE

Start the server in SSE mode:
```bash
fastmcp run server.py --transport sse --port 8000
```

Then connect your MCP client to:
```
http://localhost:8000/sse
```

For Claude Code, add to your MCP config:
```json
{
  "mcpServers": {
    "outlook": {
      "type": "sse",
      "url": "http://localhost:8000/sse"
    }
  }
}
```

---

## 5. Multi-account Setup

The server supports multiple Azure AD accounts simultaneously. Each tool accepts an optional `account` parameter to select which credentials to use.

### Environment variables

```dotenv
# Primary / default account
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=primary-secret

# Named account "work" — different tenant
ACCOUNT_WORK_TENANT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_WORK_CLIENT_ID=yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
ACCOUNT_WORK_CLIENT_SECRET=work-secret

# Named account "personal" — delegated auth, same app registration
ACCOUNT_PERSONAL_REFRESH_TOKEN=personal-refresh-token
```

> When `TENANT_ID` or `CLIENT_ID` are omitted for a named account, they fall back to the default `AZURE_*` values — so multiple users sharing the same app registration only need to supply their own token.

### Usage

```
list_accounts()
  → {"accounts": ["default", "work", "personal"], ...}

get_calendar_events(user_email="alice@company.com")
  → uses default account

get_calendar_events(user_email="bob@work.com", account="work")
  → uses work account credentials

list_calendars(user_email="me@gmail.com", account="personal")
  → uses personal delegated token
```

### Multi-account config (uvx)

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/lucassampsouza/outlook-mcp-server", "outlook-mcp-server"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "ACCOUNT_WORK_TENANT_ID": "...",
        "ACCOUNT_WORK_CLIENT_ID": "...",
        "ACCOUNT_WORK_REFRESH_TOKEN": "..."
      }
    }
  }
}
```

---

## Environment Variables Reference

| Variable | Required | Description |
|---|---|---|
| `AZURE_TENANT_ID` | Yes | Azure AD Directory (tenant) ID |
| `AZURE_CLIENT_ID` | Yes | Application (client) ID |
| `AZURE_CLIENT_SECRET` | Mode 1 | Client secret value (application auth) |
| `AZURE_REFRESH_TOKEN` | Mode 2 | Refresh token (delegated auth, set by setup.py) |
| `ACCOUNT_{NAME}_TENANT_ID` | No | Tenant ID for a named account |
| `ACCOUNT_{NAME}_CLIENT_ID` | No | Client ID for a named account |
| `ACCOUNT_{NAME}_CLIENT_SECRET` | No | Client secret for a named account (application auth) |
| `ACCOUNT_{NAME}_REFRESH_TOKEN` | No | Refresh token for a named account (delegated auth) |

---

## Requirements

- Python 3.10+
- `fastmcp >= 2.0.0`
- `httpx >= 0.27.0`
- `python-dotenv >= 1.0.0`
- `msal >= 1.28.0`
