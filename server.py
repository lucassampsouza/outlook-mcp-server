"""
Outlook Calendar MCP Server
Exposes Microsoft Graph API calendar operations as MCP tools.

Authentication uses the Client Credentials (app-only) flow.

Single-account (default) environment variables:
    AZURE_TENANT_ID     - Your Azure AD tenant ID
    AZURE_CLIENT_ID     - Application (client) ID from Azure App Registration
    AZURE_CLIENT_SECRET - Client secret from Azure App Registration

Multi-account support (named profiles):
    ACCOUNT_{NAME}_TENANT_ID     - Tenant ID for the named account
    ACCOUNT_{NAME}_CLIENT_ID     - Client ID for the named account
    ACCOUNT_{NAME}_CLIENT_SECRET - Client secret for the named account

    Example names: ACCOUNT_WORK_TENANT_ID, ACCOUNT_PERSONAL_TENANT_ID
    Use account="work" or account="personal" in any tool call.

See SETUP.md for detailed setup instructions.
"""

import os
from datetime import datetime, timedelta, timezone
from typing import Optional

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP

load_dotenv()

mcp = FastMCP("Outlook Calendar MCP Server")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

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
    if tenant and client and secret:
        accounts["default"] = {
            "tenant_id": tenant,
            "client_id": client,
            "client_secret": secret,
        }

    # Named accounts: ACCOUNT_{NAME}_TENANT_ID
    for key, value in os.environ.items():
        if key.startswith("ACCOUNT_") and key.endswith("_TENANT_ID"):
            # Extract the name part between ACCOUNT_ and _TENANT_ID
            name = key[len("ACCOUNT_"):-len("_TENANT_ID")].lower()
            prefix = f"ACCOUNT_{key[len('ACCOUNT_'):-len('_TENANT_ID')]}_"
            client_id = os.environ.get(f"{prefix}CLIENT_ID")
            client_secret = os.environ.get(f"{prefix}CLIENT_SECRET")
            if client_id and client_secret:
                accounts[name] = {
                    "tenant_id": value,
                    "client_id": client_id,
                    "client_secret": client_secret,
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
            "Check your .env file — add ACCOUNT_{NAME}_TENANT_ID / "
            "ACCOUNT_{NAME}_CLIENT_ID / ACCOUNT_{NAME}_CLIENT_SECRET "
            "for each additional account."
        )
    return _accounts[account]


def get_access_token(account: str = "default") -> str:
    """Obtain an access token for the given account, with basic caching."""
    now = datetime.now(timezone.utc)
    cached = _token_cache.get(account)
    if cached and cached["expires_at"] > now:
        return cached["value"]

    creds = _get_creds(account)
    token_url = (
        f"https://login.microsoftonline.com/{creds['tenant_id']}/oauth2/v2.0/token"
    )
    response = httpx.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": creds["client_id"],
            "client_secret": creds["client_secret"],
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()

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


if __name__ == "__main__":
    mcp.run()
