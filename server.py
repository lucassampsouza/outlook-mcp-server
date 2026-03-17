"""
Outlook Calendar MCP Server
Exposes Microsoft Graph API calendar operations as MCP tools.

Supports two authentication modes:

1. Application permissions (Client Credentials flow) — requires admin consent:
    AZURE_TENANT_ID     - Your Azure AD tenant ID
    AZURE_CLIENT_ID     - Application (client) ID from Azure App Registration
    AZURE_CLIENT_SECRET - Client secret from Azure App Registration

2. Delegated permissions (Refresh Token flow) — no admin consent required:
    AZURE_TENANT_ID      - Your Azure AD tenant ID
    AZURE_CLIENT_ID      - Application (client) ID from Azure App Registration
    AZURE_REFRESH_TOKEN  - Refresh token obtained via setup.py Device Code flow

Multi-account support (named profiles) — same two modes per account:
    ACCOUNT_{NAME}_TENANT_ID     - Tenant ID for the named account
    ACCOUNT_{NAME}_CLIENT_ID     - Client ID for the named account
    ACCOUNT_{NAME}_CLIENT_SECRET - Client secret (application auth)
      OR
    ACCOUNT_{NAME}_REFRESH_TOKEN - Refresh token (delegated auth)

    Example names: ACCOUNT_WORK_TENANT_ID, ACCOUNT_PERSONAL_TENANT_ID
    Use account="work" or account="personal" in any tool call.

See SETUP.md for detailed setup instructions.
"""

import json
import os
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP

load_dotenv(Path(__file__).parent / ".env", override=True)

mcp = FastMCP("Outlook Calendar MCP Server")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Explicit delegated scopes requested during Device Code flow.
# Using .default fails in multi-tenant scenarios where the app hasn't been
# pre-authorized in the user's tenant — explicit scopes work with any tenant
# and drive a clear user consent screen.
DELEGATED_SCOPES = (
    "https://graph.microsoft.com/Calendars.Read "
    "https://graph.microsoft.com/Calendars.Read.Shared "
    "https://graph.microsoft.com/Calendars.ReadBasic "
    "https://graph.microsoft.com/Calendars.ReadWrite "
    "https://graph.microsoft.com/Calendars.ReadWrite.Shared "
    "https://graph.microsoft.com/User.Read "
    "offline_access"
)

# ---------------------------------------------------------------------------
# Multi-account credential loading
# ---------------------------------------------------------------------------

def _load_accounts() -> dict[str, dict]:
    """
    Build a mapping of account_name -> credential dict from environment variables.

    Always attempts to load a "default" account from AZURE_TENANT_ID /
    AZURE_CLIENT_ID / AZURE_CLIENT_SECRET.  Additionally scans for any env
    vars matching ACCOUNT_{NAME}_TENANT_ID and groups them into named profiles.
    """
    accounts: dict[str, dict] = {}

    # Default / single-account credentials (backward compatible)
    tenant = os.environ.get("AZURE_TENANT_ID")
    client = os.environ.get("AZURE_CLIENT_ID")
    secret = os.environ.get("AZURE_CLIENT_SECRET")
    refresh_token = os.environ.get("AZURE_REFRESH_TOKEN")
    if tenant and client:
        if secret:
            accounts["default"] = {
                "tenant_id": tenant,
                "client_id": client,
                "client_secret": secret,
                "auth_type": "application",
            }
        elif refresh_token:
            accounts["default"] = {
                "tenant_id": tenant,
                "client_id": client,
                "refresh_token": refresh_token,
                "auth_type": "delegated",
            }

    # Named accounts.
    #
    # A named account is detected by either ACCOUNT_{NAME}_TENANT_ID or
    # ACCOUNT_{NAME}_REFRESH_TOKEN / ACCOUNT_{NAME}_CLIENT_SECRET being present.
    # CLIENT_ID and TENANT_ID fall back to the default AZURE_* values when not
    # explicitly set, so multiple users sharing the same app registration only
    # need to supply their own REFRESH_TOKEN (or CLIENT_SECRET).
    named: dict[str, str] = {}  # raw_name -> prefix
    for key in os.environ:
        for suffix in ("_TENANT_ID", "_CLIENT_ID", "_CLIENT_SECRET", "_REFRESH_TOKEN"):
            if key.startswith("ACCOUNT_") and key.endswith(suffix):
                raw = key[len("ACCOUNT_"):-len(suffix)]
                named[raw] = f"ACCOUNT_{raw}_"

    for raw, prefix in named.items():
        name = raw.lower()
        acc_tenant = os.environ.get(f"{prefix}TENANT_ID") or tenant
        acc_client = os.environ.get(f"{prefix}CLIENT_ID") or client
        if not acc_tenant or not acc_client:
            continue
        client_secret = os.environ.get(f"{prefix}CLIENT_SECRET")
        refresh_tok = os.environ.get(f"{prefix}REFRESH_TOKEN")
        if client_secret:
            accounts[name] = {
                "tenant_id": acc_tenant,
                "client_id": acc_client,
                "client_secret": client_secret,
                "auth_type": "application",
            }
        elif refresh_tok:
            accounts[name] = {
                "tenant_id": acc_tenant,
                "client_id": acc_client,
                "refresh_token": refresh_tok,
                "auth_type": "delegated",
            }

    return accounts


_accounts: dict[str, dict] = _load_accounts()

# Token cache keyed by account name
_token_cache: dict[str, dict] = {}


def _get_creds(account: str) -> dict:
    """Return credentials for the named account, raising clearly if not found."""
    if account not in _accounts:
        configured = list(_accounts.keys()) or ["(none)"]
        raise ValueError(
            f"Unknown account '{account}'. "
            f"Configured accounts: {configured}. "
            "Check your .env file — for application auth add "
            "ACCOUNT_{NAME}_TENANT_ID / ACCOUNT_{NAME}_CLIENT_ID / ACCOUNT_{NAME}_CLIENT_SECRET; "
            "for delegated auth add "
            "ACCOUNT_{NAME}_TENANT_ID / ACCOUNT_{NAME}_CLIENT_ID / ACCOUNT_{NAME}_REFRESH_TOKEN. "
            "Run setup.py to generate credentials."
        )
    return _accounts[account]


def get_access_token(account: str = "default") -> str:
    """Obtain an access token for the given account, with basic caching.

    Supports two authentication flows determined by account configuration:
    - Application (client_credentials): uses client_secret, requires admin consent.
    - Delegated (refresh_token): uses a stored refresh token, no admin consent needed.
    """
    now = datetime.now(timezone.utc)
    cached = _token_cache.get(account)
    if cached and cached["expires_at"] > now:
        return cached["value"]

    creds = _get_creds(account)
    token_url = (
        f"https://login.microsoftonline.com/{creds['tenant_id']}/oauth2/v2.0/token"
    )

    if creds.get("auth_type") == "delegated":
        token_data = {
            "grant_type": "refresh_token",
            "client_id": creds["client_id"],
            "refresh_token": creds["refresh_token"],
            "scope": DELEGATED_SCOPES,
        }
    else:
        token_data = {
            "grant_type": "client_credentials",
            "client_id": creds["client_id"],
            "client_secret": creds["client_secret"],
            "scope": "https://graph.microsoft.com/.default",
        }

    response = httpx.post(token_url, data=token_data, timeout=30)
    response.raise_for_status()
    data = response.json()

    # Microsoft rotates refresh tokens on every use — the old one is immediately
    # revoked.  Update both in-memory and .env so a process restart doesn't
    # leave a stale (revoked) token on disk.
    if "refresh_token" in data and creds.get("auth_type") == "delegated":
        new_rt = data["refresh_token"]
        _accounts[account]["refresh_token"] = new_rt
        _save_account_to_env(account, creds["tenant_id"], creds["client_id"], new_rt)

    expires_in = int(data.get("expires_in", 3600))
    _token_cache[account] = {
        "value": data["access_token"],
        "expires_at": now + timedelta(seconds=expires_in - 60),
    }
    return data["access_token"]


def _graph(method: str, path: str, account: str = "default", **kwargs) -> dict:
    """Make an authenticated request to the Microsoft Graph API."""
    token = get_access_token(account)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    response = httpx.request(
        method,
        f"{GRAPH_BASE_URL}{path}",
        headers=headers,
        timeout=30,
        **kwargs,
    )
    response.raise_for_status()
    if response.status_code == 204:
        return {}
    return response.json()


# ---------------------------------------------------------------------------
# Device Code auth helpers
# ---------------------------------------------------------------------------

# In-memory cache (same process). Cross-process state is stored on disk.
_pending_device_flows: dict[str, dict] = {}

_FLOWS_FILE = Path(__file__).parent / ".pending_flows.json"

DEVICE_CODE_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"


def _flows_load() -> dict:
    """Read all pending flows from disk."""
    if _FLOWS_FILE.exists():
        try:
            return json.loads(_FLOWS_FILE.read_text())
        except Exception:
            pass
    return {}


def _flows_save(flows: dict) -> None:
    """Persist all pending flows to disk."""
    _FLOWS_FILE.write_text(json.dumps(flows, indent=2))


def _flows_delete(account_name: str) -> None:
    """Remove a single flow from disk."""
    flows = _flows_load()
    flows.pop(account_name, None)
    if flows:
        _flows_save(flows)
    elif _FLOWS_FILE.exists():
        _FLOWS_FILE.unlink()


def _save_account_to_env(
    account_name: str, tenant_id: str, client_id: str, refresh_token: str
) -> None:
    """Persist delegated credentials to .env, replacing any previous entry."""
    env_file = Path(__file__).parent / ".env"

    if account_name == "default":
        keys_to_remove = {
            "AZURE_TENANT_ID", "AZURE_CLIENT_ID",
            "AZURE_CLIENT_SECRET", "AZURE_REFRESH_TOKEN",
        }
        new_lines = [
            f'AZURE_TENANT_ID="{tenant_id}"',
            f'AZURE_CLIENT_ID="{client_id}"',
            f'AZURE_REFRESH_TOKEN="{refresh_token}"',
        ]
    else:
        prefix = f"ACCOUNT_{account_name.upper()}_"
        keys_to_remove = {
            f"{prefix}TENANT_ID", f"{prefix}CLIENT_ID",
            f"{prefix}CLIENT_SECRET", f"{prefix}REFRESH_TOKEN",
        }
        new_lines = [
            f'{prefix}TENANT_ID="{tenant_id}"',
            f'{prefix}CLIENT_ID="{client_id}"',
            f'{prefix}REFRESH_TOKEN="{refresh_token}"',
        ]

    existing: list[str] = []
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            key = line.split("=")[0].strip() if "=" in line else ""
            if key not in keys_to_remove:
                existing.append(line)

    env_file.write_text("\n".join(existing + new_lines) + "\n")


def _reload_accounts() -> None:
    """Re-read .env and refresh the in-memory accounts dict."""
    load_dotenv(Path(__file__).parent / ".env", override=True)
    _accounts.clear()
    _accounts.update(_load_accounts())


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------


@mcp.tool()
def list_accounts() -> dict:
    """List all configured Azure AD accounts available to this server.

    Use the returned account names as the `account` parameter in other tools.
    """
    return {
        "accounts": list(_accounts.keys()),
        "note": (
            "Use one of these names as the `account` parameter in any tool. "
            "Omit `account` (or use 'default') to use the primary account."
        ),
    }


@mcp.tool()
def start_device_code_auth(
    account_name: str = "default",
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
) -> dict:
    """Begin a Device Code authentication flow to add a delegated account.

    Use this to authenticate a Microsoft user without needing admin consent or
    a client secret.  Returns a short code and URL the user must visit to sign
    in.  After sign-in, call complete_device_code_auth with the same
    account_name to finish and persist the credentials.

    Args:
        account_name: Name to assign to this account (e.g. 'alice', 'work').
            Use 'default' to replace the primary account.  Defaults to
            'default'.
        tenant_id: Azure AD tenant ID to restrict authentication to a specific
            tenant.  Defaults to 'common', which accepts users from any Azure
            AD tenant (requires the app registration to be multi-tenant).
        client_id: Azure Application (client) ID.  Falls back to
            AZURE_CLIENT_ID when omitted.
    """
    # 'common' allows users from any tenant; a specific tenant_id restricts
    # authentication to that tenant only.
    resolved_tenant = tenant_id or "common"
    resolved_client = client_id or os.environ.get("AZURE_CLIENT_ID")

    if not resolved_tenant or not resolved_client:
        return {
            "error": (
                "tenant_id and client_id are required. "
                "Pass them as arguments or set AZURE_TENANT_ID / AZURE_CLIENT_ID in .env."
            )
        }

    response = httpx.post(
        DEVICE_CODE_URL.format(tenant_id=resolved_tenant),
        data={
            "client_id": resolved_client,
            "scope": DELEGATED_SCOPES,
        },
        timeout=30,
    )
    if not response.is_success:
        body = response.json() if response.headers.get("content-type", "").startswith("application/json") else response.text
        return {
            "error": f"HTTP {response.status_code}",
            "microsoft_error": body,
            "hint": (
                "Common causes: (1) wrong tenant_id or client_id; "
                "(2) 'Allow public client flows' not enabled in the Azure App Registration "
                "(Portal → App Registration → Authentication → Advanced settings); "
                "(3) app does not exist in this tenant."
            ),
        }
    flow = response.json()

    flow_state = {
        "tenant_id": resolved_tenant,
        "client_id": resolved_client,
        "device_code": flow["device_code"],
        "interval": int(flow.get("interval", 5)),
        "expires_in": int(flow.get("expires_in", 900)),
    }
    # Persist to disk so complete_device_code_auth works across processes/sessions.
    _pending_device_flows[account_name] = flow_state
    flows = _flows_load()
    flows[account_name] = flow_state
    _flows_save(flows)

    return {
        "user_code": flow["user_code"],
        "verification_uri": flow["verification_uri"],
        "message": flow.get("message"),
        "next_step": (
            f"After signing in, call complete_device_code_auth "
            f"with account_name='{account_name}'."
        ),
    }


@mcp.tool()
def get_admin_consent_url(
    tenant: str,
    client_id: Optional[str] = None,
) -> dict:
    """Generate an admin consent URL for a tenant whose policy blocks user consent.

    Send this URL to the Azure AD admin of the target organization.  Once they
    approve it, all users in that tenant can authenticate without needing
    per-user consent.

    Args:
        tenant: The tenant domain (e.g. 'contoso.com') or tenant ID (UUID) of
            the organization whose admin needs to approve the app.
        client_id: Azure Application (client) ID.  Falls back to
            AZURE_CLIENT_ID when omitted.
    """
    resolved_client = client_id or os.environ.get("AZURE_CLIENT_ID")
    if not resolved_client:
        return {"error": "client_id is required or set AZURE_CLIENT_ID in .env."}

    url = (
        f"https://login.microsoftonline.com/{tenant}/adminconsent"
        f"?client_id={resolved_client}"
    )
    return {
        "admin_consent_url": url,
        "instructions": (
            f"Send this URL to an Azure AD Global Administrator of '{tenant}'. "
            "They must sign in with an admin account and click Accept. "
            "After approval, users in that tenant can authenticate normally."
        ),
    }


@mcp.tool()
def complete_device_code_auth(account_name: str = "default") -> dict:
    """Poll Microsoft to finish a pending Device Code authentication.

    Call this after the user has completed sign-in at the URL shown by
    start_device_code_auth.  Polls until authentication succeeds or the
    flow expires, then saves the refresh token to .env and activates the
    account immediately — no server restart required.

    Args:
        account_name: The same name used in start_device_code_auth.
    """
    # Check in-memory first; fall back to disk (cross-process / cross-session).
    flow = _pending_device_flows.get(account_name) or _flows_load().get(account_name)
    if not flow:
        return {
            "error": (
                f"No pending authentication for account '{account_name}'. "
                "Call start_device_code_auth first."
            )
        }

    token_url = TOKEN_URL.format(tenant_id=flow["tenant_id"])
    interval = flow["interval"]
    deadline = time.monotonic() + flow["expires_in"]

    while time.monotonic() < deadline:
        time.sleep(interval)
        resp = httpx.post(
            token_url,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": flow["client_id"],
                "device_code": flow["device_code"],
            },
            timeout=30,
        )
        data = resp.json()

        if "access_token" in data:
            refresh_token = data.get("refresh_token", "")
            _save_account_to_env(
                account_name, flow["tenant_id"], flow["client_id"], refresh_token
            )
            _reload_accounts()
            _pending_device_flows.pop(account_name, None)
            _flows_delete(account_name)
            return {
                "status": "success",
                "account": account_name,
                "message": (
                    f"Account '{account_name}' authenticated and saved. "
                    "It is now available for all tools."
                ),
            }

        error = data.get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        # Any other error (e.g. expired_token, access_denied) is terminal
        _pending_device_flows.pop(account_name, None)
        _flows_delete(account_name)
        return {
            "error": error,
            "error_description": data.get("error_description"),
        }

    _pending_device_flows.pop(account_name, None)
    _flows_delete(account_name)
    return {
        "error": "timeout",
        "message": (
            "The device code expired before authentication completed. "
            "Call start_device_code_auth again to restart."
        ),
    }


@mcp.tool()
def list_calendars(user_email: str, account: str = "default") -> dict:
    """List all calendars for a user.

    Args:
        user_email: The user's Outlook / Microsoft 365 email address.
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    return _graph("GET", f"/users/{user_email}/calendars", account=account)


@mcp.tool()
def get_calendar_events(
    user_email: str,
    calendar_id: str = "primary",
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    top: int = 10,
    account: str = "default",
) -> dict:
    """Retrieve events from a calendar within a time window.

    Args:
        user_email: The user's Outlook / Microsoft 365 email address.
        calendar_id: Calendar ID returned by list_calendars, or 'primary' for
            the default calendar.
        start_datetime: Start of range in ISO 8601 UTC format
            (e.g. '2024-06-01T00:00:00Z'). Defaults to now.
        end_datetime: End of range in ISO 8601 UTC format. Defaults to 7 days
            from now.
        top: Maximum number of events to return (1–50, default 10).
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    now = datetime.now(timezone.utc)
    if start_datetime is None:
        start_datetime = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    if end_datetime is None:
        end_datetime = (now + timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")

    path = (
        f"/users/{user_email}/calendar/events"
        if calendar_id == "primary"
        else f"/users/{user_email}/calendars/{calendar_id}/events"
    )

    params = {
        "$top": max(1, min(top, 50)),
        "$filter": (
            f"start/dateTime ge '{start_datetime}' and "
            f"start/dateTime le '{end_datetime}'"
        ),
        "$orderby": "start/dateTime",
        "$select": (
            "id,subject,start,end,location,organizer,"
            "attendees,bodyPreview,isAllDay,isCancelled,webLink"
        ),
    }
    return _graph("GET", path, account=account, params=params)


@mcp.tool()
def get_event(user_email: str, event_id: str, account: str = "default") -> dict:
    """Get full details for a specific calendar event.

    Args:
        user_email: The user's email address.
        event_id: The event ID (from get_calendar_events).
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    return _graph("GET", f"/users/{user_email}/events/{event_id}", account=account)


@mcp.tool()
def create_event(
    user_email: str,
    subject: str,
    start_datetime: str,
    end_datetime: str,
    body: Optional[str] = None,
    location: Optional[str] = None,
    attendees: Optional[list[str]] = None,
    calendar_id: Optional[str] = None,
    is_online_meeting: bool = False,
    account: str = "default",
) -> dict:
    """Create a new calendar event.

    Args:
        user_email: The organiser's email address.
        subject: Event title.
        start_datetime: Start time in ISO 8601 UTC (e.g. '2024-06-15T14:00:00Z').
        end_datetime: End time in ISO 8601 UTC.
        body: Optional plain-text description.
        location: Optional location display name.
        attendees: Optional list of attendee email addresses.
        calendar_id: Target calendar ID; omit to use the default calendar.
        is_online_meeting: Set True to attach a Teams meeting link.
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    event_data: dict = {
        "subject": subject,
        "start": {"dateTime": start_datetime, "timeZone": "UTC"},
        "end": {"dateTime": end_datetime, "timeZone": "UTC"},
        "isOnlineMeeting": is_online_meeting,
    }
    if body:
        event_data["body"] = {"contentType": "text", "content": body}
    if location:
        event_data["location"] = {"displayName": location}
    if attendees:
        event_data["attendees"] = [
            {"emailAddress": {"address": email}, "type": "required"}
            for email in attendees
        ]

    path = (
        f"/users/{user_email}/calendars/{calendar_id}/events"
        if calendar_id
        else f"/users/{user_email}/events"
    )
    return _graph("POST", path, account=account, json=event_data)


@mcp.tool()
def update_event(
    user_email: str,
    event_id: str,
    subject: Optional[str] = None,
    start_datetime: Optional[str] = None,
    end_datetime: Optional[str] = None,
    body: Optional[str] = None,
    location: Optional[str] = None,
    account: str = "default",
) -> dict:
    """Update fields on an existing calendar event.

    Args:
        user_email: The user's email address.
        event_id: The event ID to update.
        subject: New event title.
        start_datetime: New start time in ISO 8601 UTC.
        end_datetime: New end time in ISO 8601 UTC.
        body: New plain-text description.
        location: New location display name.
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    event_data: dict = {}
    if subject is not None:
        event_data["subject"] = subject
    if start_datetime is not None:
        event_data["start"] = {"dateTime": start_datetime, "timeZone": "UTC"}
    if end_datetime is not None:
        event_data["end"] = {"dateTime": end_datetime, "timeZone": "UTC"}
    if body is not None:
        event_data["body"] = {"contentType": "text", "content": body}
    if location is not None:
        event_data["location"] = {"displayName": location}

    return _graph(
        "PATCH",
        f"/users/{user_email}/events/{event_id}",
        account=account,
        json=event_data,
    )


@mcp.tool()
def delete_event(user_email: str, event_id: str, account: str = "default") -> dict:
    """Permanently delete a calendar event.

    Args:
        user_email: The user's email address.
        event_id: The event ID to delete.
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    _graph("DELETE", f"/users/{user_email}/events/{event_id}", account=account)
    return {"status": "deleted", "event_id": event_id}


@mcp.tool()
def get_free_busy(
    user_emails: list[str],
    start_datetime: str,
    end_datetime: str,
    interval_minutes: int = 30,
    account: str = "default",
) -> dict:
    """Return the free/busy schedule for one or more users.

    Args:
        user_emails: List of email addresses to check.
        start_datetime: Start of the window in ISO 8601 UTC.
        end_datetime: End of the window in ISO 8601 UTC.
        interval_minutes: Granularity of the availability view in minutes
            (default 30, minimum 5).
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    payload = {
        "schedules": user_emails,
        "startTime": {"dateTime": start_datetime, "timeZone": "UTC"},
        "endTime": {"dateTime": end_datetime, "timeZone": "UTC"},
        "availabilityViewInterval": max(5, interval_minutes),
    }
    return _graph(
        "POST",
        f"/users/{user_emails[0]}/calendar/getSchedule",
        account=account,
        json=payload,
    )


@mcp.tool()
def search_events(
    user_email: str,
    query: str,
    top: int = 10,
    account: str = "default",
) -> dict:
    """Search calendar events by keyword using the Graph search endpoint.

    Args:
        user_email: The user's email address.
        query: Search keywords (subject, body, location, etc.).
        top: Maximum number of results (1–25, default 10).
        account: Which configured Azure AD account to authenticate with
            (see list_accounts). Defaults to 'default'.
    """
    params = {
        "$search": f'"{query}"',
        "$top": max(1, min(top, 25)),
        "$select": "id,subject,start,end,location,organizer,bodyPreview",
    }
    return _graph("GET", f"/users/{user_email}/events", account=account, params=params)


def main() -> None:
    import sys
    try:
        # fastmcp 2.x stores tools in _tools dict
        tools = list(mcp._tools.keys())
        print(f"[outlook-mcp] registered tools ({len(tools)}): {tools}", file=sys.stderr)
    except AttributeError:
        try:
            attrs = [a for a in dir(mcp) if "tool" in a.lower()]
            print(f"[outlook-mcp] tool-related attrs: {attrs}", file=sys.stderr)
        except Exception as e2:
            print(f"[outlook-mcp] debug failed: {e2}", file=sys.stderr)
    mcp.run()


if __name__ == "__main__":
    main()
