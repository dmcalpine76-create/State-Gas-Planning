"""
core/graph.py — Microsoft Graph access shared by every tool.

Consolidates the four diverged copies of auth, folder discovery, and email
fetching, with these fixes over the originals:

  • ONE scope set (superset) for the shared token cache — no more silent-auth
    failures when tools requesting different scopes shared one cache.       (B9)
  • fetch_folder_emails follows @odata.nextLink, so folders with more than
    one page of recent mail no longer silently drop messages.               (B7)
  • Bodies are requested in the folder list call itself (Prefer:
    outlook.body-content-type="text"), replacing one HTTP round-trip per
    email with one per folder page — scans are roughly 10x faster.          (P1)
  • Folder discovery results are cached for FOLDER_CACHE_TTL_DAYS.          (P2)
"""

import re
import json
import datetime

import msal
import requests

from .config import (TOKEN_CACHE_FILE, FOLDER_CACHE_FILE,
                     FOLDER_CACHE_TTL_DAYS)
import os

GRAPH_BASE      = "https://graph.microsoft.com/v1.0"
REQUEST_TIMEOUT = 20

# Superset of every scope any tool needs. One consent, one cache, no drift.
# Calendars.ReadWrite was added when outlook_scheduler.py moved onto this
# shared token — if you authorised before that, delete
# .outlook_token_cache.bin and re-run setup, or calendar writes will 403.
SCOPES = ["Mail.ReadWrite", "Tasks.ReadWrite", "Calendars.ReadWrite", "User.Read"]

SKIP_SYSTEM = {
    "deleteditems", "junkemail", "drafts", "outbox", "archive",
    "conversationhistory", "syncissues", "recoverableitemsdeletions",
    "recoverableitemsroot", "recoverableitemspurges",
    "recoverableitemsversions", "scheduledactions", "searchfolders",
    "spam", "junk",
}


# ─────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────

def _client_id() -> str:
    return os.environ.get("OUTLOOK_CLIENT_ID", "")

def _authority() -> str:
    return ("https://login.microsoftonline.com/"
            + os.environ.get("OUTLOOK_TENANT_ID", "consumers"))

def _load_cache():
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())
        try:
            TOKEN_CACHE_FILE.chmod(0o600)
        except Exception:
            pass

def _build_app(cache):
    if not _client_id():
        raise RuntimeError("OUTLOOK_CLIENT_ID not set. Add it to your .env file.")
    return msal.PublicClientApplication(
        _client_id(), authority=_authority(), token_cache=cache)


def get_access_token() -> str:
    cache    = _load_cache()
    app      = _build_app(cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    raise RuntimeError(
        "Outlook token missing or expired.\n"
        "Run:  py inbox_actions.py setup   (any tool's setup works — auth is shared)"
    )


def setup_auth():
    """Interactive device-code auth flow. Re-run if scopes change."""
    if not _client_id():
        print("\n❌  OUTLOOK_CLIENT_ID not found in environment / .env")
        return
    cache = _load_cache()
    app   = _build_app(cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            print(f"\n✅  Already authenticated as: {accounts[0]['username']}")
            print("    (If scopes changed, delete .outlook_token_cache.bin and re-run setup)\n")
            return

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"\n❌  Failed: {json.dumps(flow, indent=2)}\n")
        return

    print("\n" + "─" * 60)
    print("  OUTLOOK AUTHORISATION")
    print("─" * 60)
    print(f"\n  1. Open:  https://microsoft.com/devicelogin")
    print(f"  2. Enter: {flow['user_code']}")
    print(f"  3. Sign in with your Microsoft account")
    print(f"     (Grant: {', '.join(SCOPES)})\n")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _save_cache(cache)
        try:
            me = requests.get(
                f"{GRAPH_BASE}/me",
                headers={"Authorization": f"Bearer {result['access_token']}"},
                timeout=10,
            ).json()
            print(f"✅  Authenticated: {me.get('displayName')} "
                  f"({me.get('mail') or me.get('userPrincipalName')})\n")
        except Exception:
            print("✅  Authenticated successfully!\n")
    else:
        print(f"\n❌  {result.get('error_description', result)}\n")


# ─────────────────────────────────────────────
# LOW-LEVEL HELPERS
# ─────────────────────────────────────────────

def graph_get(token: str, url: str, params: dict = None,
              headers: dict = None) -> dict:
    if not url.startswith("http"):
        url = f"{GRAPH_BASE}{url}"
    hdrs = {"Authorization": f"Bearer {token}"}
    if headers:
        hdrs.update(headers)
    resp = requests.get(url, headers=hdrs, params=params or {},
                        timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def graph_post(token: str, url: str, body: dict) -> dict:
    if not url.startswith("http"):
        url = f"{GRAPH_BASE}{url}"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}",
                 "Content-Type":  "application/json"},
        json=body,
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    return resp.json()


def graph_patch(token: str, url: str, body: dict) -> dict:
    if not url.startswith("http"):
        url = f"{GRAPH_BASE}{url}"
    resp = requests.patch(
        url,
        headers={"Authorization": f"Bearer {token}",
                 "Content-Type":  "application/json"},
        json=body,
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    # PATCH can legitimately return 204 No Content.
    if not resp.content:
        return {}
    return resp.json()


def has_calendar_scope() -> bool:
    """
    True if the cached token was issued with Calendars.ReadWrite. Lets the
    scheduler fail with a useful message instead of a bare 403 when the cache
    predates the scope being added.
    """
    try:
        cache    = _load_cache()
        app      = _build_app(cache)
        accounts = app.get_accounts()
        if not accounts:
            return False
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            return False
        granted = " ".join(result.get("scope", "") or "").lower() \
            if isinstance(result.get("scope"), list) else str(result.get("scope", "")).lower()
        return "calendars.readwrite" in granted
    except Exception:
        return False


def strip_html(body: str) -> str:
    body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.DOTALL)
    body = re.sub(r'<[^>]+>', ' ', body)
    body = re.sub(r'[ \t]+', ' ', body)
    return re.sub(r'\n{3,}', '\n\n', body).strip()


# ─────────────────────────────────────────────
# FOLDER DISCOVERY  (BFS, cached)
# ─────────────────────────────────────────────

def _fetch_children(token: str, parent_id: str) -> list:
    children = []
    url      = f"{GRAPH_BASE}/me/mailFolders/{parent_id}/childFolders"
    params   = {"$top": 100, "$select": "id,displayName"}
    while url:
        try:
            resp = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                                params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code == 404:
                break
            resp.raise_for_status()
            data = resp.json()
            children.extend(data.get("value", []))
            url    = data.get("@odata.nextLink")
            params = {}
        except Exception:
            break
    return children


def _load_folder_cache():
    if not FOLDER_CACHE_FILE.exists():
        return None
    try:
        cached = json.loads(FOLDER_CACHE_FILE.read_text(encoding="utf-8"))
        age = (datetime.date.today()
               - datetime.date.fromisoformat(cached.get("cached", "2000-01-01"))).days
        if age <= FOLDER_CACHE_TTL_DAYS:
            return [tuple(f) for f in cached.get("folders", [])]
    except Exception:
        pass
    return None


def get_all_folders(token: str, refresh: bool = False) -> list:
    """
    BFS discovery of every non-system folder → list of (folder_id, path).
    Results are cached for FOLDER_CACHE_TTL_DAYS; pass refresh=True (or the
    --refresh-folders flag on any tool) to force rediscovery.
    """
    if not refresh:
        cached = _load_folder_cache()
        if cached:
            print(f"     📁  {len(cached)} folders (cached — "
                  f"use --refresh-folders to rediscover)")
            return cached

    queue, results, seen = [], [], set()

    def enqueue(fid, fpath):
        if fid not in seen:
            seen.add(fid)
            queue.append((fid, fpath))
            results.append((fid, fpath))

    enqueue("inbox",     "Inbox")
    enqueue("sentitems", "Sent Items")

    print("     Discovering top-level folders…")
    url    = f"{GRAPH_BASE}/me/mailFolders"
    params = {"$top": 100, "$select": "id,displayName"}
    while url:
        try:
            data = graph_get(token, url, params)
            for f in data.get("value", []):
                fname = f.get("displayName", "")
                if fname.lower().replace(" ", "") not in SKIP_SYSTEM:
                    enqueue(f["id"], fname)
            url    = data.get("@odata.nextLink")
            params = {}
        except Exception as e:
            print(f"   ⚠️  Top-level folder error: {e}")
            break

    print("     Recursing into subfolders…")
    processed = 0
    while queue:
        parent_id, parent_path = queue.pop(0)
        for child in _fetch_children(token, parent_id):
            cfname = child.get("displayName", "")
            if cfname.lower().replace(" ", "") not in SKIP_SYSTEM:
                enqueue(child["id"], f"{parent_path} > {cfname}")
        processed += 1
        if processed % 20 == 0:
            print(f"     … {processed} folders processed, {len(results)} found")

    print(f"     ✅  {len(results)} folders discovered")
    try:
        FOLDER_CACHE_FILE.write_text(json.dumps({
            "cached":  datetime.date.today().isoformat(),
            "folders": results,
        }), encoding="utf-8")
    except Exception:
        pass
    return results


# ─────────────────────────────────────────────
# EMAIL FETCHING
# ─────────────────────────────────────────────

def _parse_msg(msg: dict, folder_path: str, body_chars: int) -> dict:
    sender  = msg.get("from", {}).get("emailAddress", {})
    to_list = [
        r.get("emailAddress", {}).get("name", "") or r.get("emailAddress", {}).get("address", "")
        for r in msg.get("toRecipients", [])[:5]
    ]
    ts = msg.get("receivedDateTime") or msg.get("sentDateTime", "")

    # Body arrives with the list response (P1). With the Prefer header Graph
    # returns text; fall back to stripping HTML if it didn't.
    body_obj = msg.get("body", {}) or {}
    body     = body_obj.get("content", "") or ""
    if body_obj.get("contentType") == "html":
        body = strip_html(body)
    body = body[:body_chars]

    return {
        "id":          msg.get("id", ""),
        "subject":     msg.get("subject", "(no subject)").strip(),
        "from_name":   sender.get("name", ""),
        "from_email":  sender.get("address", ""),
        "to":          ", ".join(to_list),
        "datetime":    ts[:19].replace("T", " ") if ts else "",
        "preview":     msg.get("bodyPreview", "")[:300].strip(),
        "importance":  msg.get("importance", "normal"),
        "is_read":     msg.get("isRead", True),
        "has_attach":  msg.get("hasAttachments", False),
        "folder":      folder_path,
        "body":        body,
    }


def fetch_folder_emails(token: str, folder_id: str, folder_path: str,
                        since: datetime.datetime, max_per_folder: int,
                        body_chars: int) -> list:
    """
    Fetch emails in a folder received/sent on or after `since`, following
    pagination up to max_per_folder (B7). Bodies come back in the same
    request (P1).
    """
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    results   = []

    for date_field in ("receivedDateTime", "sentDateTime"):
        url    = f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
        params = {
            "$filter":  f"{date_field} ge {since_iso}",
            "$select":  "id,subject,from,toRecipients,receivedDateTime,sentDateTime,"
                        "hasAttachments,importance,isRead,bodyPreview,body",
            "$orderby": f"{date_field} desc",
            "$top":     min(max_per_folder, 50),
        }
        found_any = False
        while url and len(results) < max_per_folder:
            try:
                resp = requests.get(
                    url,
                    headers={"Authorization": f"Bearer {token}",
                             "Prefer": 'outlook.body-content-type="text"'},
                    params=params,
                    timeout=REQUEST_TIMEOUT,
                )
                if resp.status_code in (400, 404, 501):
                    break
                resp.raise_for_status()
                data = resp.json()
                for msg in data.get("value", []):
                    mid = msg.get("id", "")
                    if mid and not any(e["id"] == mid for e in results):
                        results.append(_parse_msg(msg, folder_path, body_chars))
                        found_any = True
                url    = data.get("@odata.nextLink")
                params = {}
            except Exception:
                break
        if found_any:
            break  # this date field worked; don't double-fetch with the other
    return results[:max_per_folder]


def run_scan(token: str, lookback_days: int, *, max_per_folder: int,
             max_total: int, body_chars: int, label: str = "EMAIL SCAN",
             refresh_folders: bool = False) -> list:
    """Discover folders, fetch recent emails (bodies included), deduplicate."""
    since    = (datetime.datetime.now(datetime.timezone.utc)
                - datetime.timedelta(days=lookback_days))
    since_dt = since.replace(hour=0, minute=0, second=0, microsecond=0)

    print(f"\n{'─'*60}")
    print(f"  {label}")
    print(f"{'─'*60}")
    print(f"  Lookback:  {lookback_days} days  "
          f"({since_dt.strftime('%d %B %Y')} → today)")
    print(f"{'─'*60}\n")

    print("  📁  Step 1: Folder discovery")
    folders = get_all_folders(token, refresh=refresh_folders)
    print(f"       {len(folders)} folders to scan\n")

    print("  📬  Step 2: Fetching emails (bodies included)…")
    all_msgs = {}
    for folder_id, folder_path in folders:
        msgs = fetch_folder_emails(token, folder_id, folder_path, since_dt,
                                   max_per_folder, body_chars)
        new = 0
        for m in msgs:
            if m["id"] not in all_msgs:
                all_msgs[m["id"]] = m
                new += 1
        if new:
            print(f"     [{folder_path}]  {new} emails")
        if len(all_msgs) >= max_total:
            print(f"\n  ⚠️  Cap of {max_total} emails reached — stopping scan")
            break

    emails = sorted(all_msgs.values(), key=lambda e: e["datetime"], reverse=True)
    print(f"\n  ✅  {len(emails)} unique emails in the last {lookback_days} days\n")
    return emails


# ─────────────────────────────────────────────
# EMAIL → PROMPT CONTEXT
# ─────────────────────────────────────────────

def build_email_context(emails: list, lookback_days: int = None) -> str:
    """Render parsed emails as the plain-text block used in every prompt."""
    lines = [f"TOTAL EMAILS: {len(emails)}"]
    if lookback_days is not None:
        today = datetime.date.today()
        since = today - datetime.timedelta(days=lookback_days)
        lines.append(f"DATE RANGE: {since.strftime('%d %B %Y')} "
                     f"to {today.strftime('%d %B %Y')}")
    lines.append("")
    for i, e in enumerate(emails):
        lines += [
            f"{'='*60}",
            f"EMAIL {i+1}",
            f"Date:       {e['datetime']}",
            f"From:       {e['from_name']} <{e['from_email']}>",
            f"To:         {e['to']}",
            f"Subject:    {e['subject']}",
            f"Folder:     {e['folder']}",
            f"Unread:     {'YES' if not e.get('is_read', True) else 'no'}",
        ]
        if e.get("has_attach"):
            lines.append("Attachments: YES")
        if e.get("importance") == "high":
            lines.append("Importance:  HIGH")
        lines += [f"\nBODY:\n{e.get('body', '[no body]')}", ""]
    return "\n".join(lines)


def top_senders_block(emails: list, n: int = 15) -> str:
    """'TOP SENDERS' summary used by the post-run knowledge patch prompts."""
    counts: dict = {}
    for e in emails:
        key = (e.get("from_name", ""), e.get("from_email", ""))
        counts[key] = counts.get(key, 0) + 1
    top   = sorted(counts.items(), key=lambda x: -x[1])[:n]
    lines = ["TOP SENDERS THIS WINDOW:"]
    for (name, addr), count in top:
        lines.append(f"  {name} <{addr}>  ({count} emails)")
    return "\n".join(lines)
