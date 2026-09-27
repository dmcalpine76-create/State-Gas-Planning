"""
matter_log.py  —  State Gas Comprehensive Matter Log Generator
---------------------------------------------------------------
A standalone companion to board_briefing.py that produces a detailed,
uncompressed narrative log of every board-relevant matter from the
review period. Nothing in this script touches board_report_data.json
or the PowerPoint pipeline — it writes only to its own output file.

Purpose
───────
board_briefing.py is optimised to produce 4-6 sentence slide briefings
for the PowerPoint template. That compression is intentional for the
board deck but loses important detail. This script preserves that detail:
  • One properly written paragraph per matter (8-12 sentences)
  • Written in Doug McAlpine's first-person voice
  • Grouped by strategic area, ordered by significance
  • Designed for Doug's own review, as working notes, and as source
    material when editing the board deck content

The output is a single Markdown file:
  StateGas_MatterLog_YYYYMMDD.md

It does NOT write to board_report_data.json and does NOT affect the
PowerPoint generator.

Usage
─────
  py matter_log.py setup           — one-time browser auth (shared with board_briefing.py)
  py matter_log.py run             — scan last 60 days and generate log
  py matter_log.py run --days 30   — custom lookback window
  py matter_log.py test            — verify connection

The script shares .env and .outlook_token_cache.bin with board_briefing.py.
Place both scripts in the same folder.

Prerequisites
─────────────
  pip install msal requests anthropic python-dotenv
"""

import os
import sys
import re
import json
import time
import datetime
import requests
import msal
import anthropic
import argparse
from pathlib import Path
from dotenv import load_dotenv, find_dotenv

# Optional: load persistent context for richer analysis
try:
    from board_briefing_context import BoardBriefingContext
    _CTX_AVAILABLE = True
except ImportError:
    _CTX_AVAILABLE = False


# ─────────────────────────────────────────────
# .ENV LOADING
# ─────────────────────────────────────────────

_here   = Path(__file__).resolve().parent
_loaded = False
for _candidate in [_here, _here.parent, _here.parent.parent]:
    _env = _candidate / ".env"
    if _env.exists():
        load_dotenv(_env)
        _loaded = True
        break
if not _loaded:
    _found = find_dotenv(usecwd=True)
    if _found:
        load_dotenv(_found)


# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

CLIENT_ID  = os.environ.get("OUTLOOK_CLIENT_ID", "")
AUTHORITY  = "https://login.microsoftonline.com/" + os.environ.get("OUTLOOK_TENANT_ID", "consumers")
SCOPES     = ["Mail.Read", "User.Read"]
CACHE_FILE = _here / ".outlook_token_cache.bin"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

DEFAULT_LOOKBACK_DAYS  = 60
MAX_EMAILS_PER_FOLDER  = 100
MAX_EMAILS_TOTAL       = 600
MAX_BODY_CHARS         = 1500    # matched to board_briefing.py — keeps batches within token limit
MAX_EMAILS_PER_BATCH   = 80      # reduced: matter_log prompt is larger than board_briefing, needs more headroom
REQUEST_TIMEOUT        = 20

# Token budget
_CHARS_PER_TOKEN  = 3.5
_MAX_BATCH_TOKENS = 155000

SKIP_SYSTEM = {
    "deleteditems", "junkemail", "drafts", "outbox", "archive",
    "conversationhistory", "syncissues", "recoverableitemsdeletions",
    "recoverableitemsroot", "recoverableitemspurges",
    "recoverableitemsversions", "scheduledactions", "searchfolders",
    "spam", "junk",
}

COMPANY_NAME   = "State Gas Limited"
COMPANY_TICKER = "ASX: GAS"

OUTPUT_DIR = _here

# ─── Strategic areas for grouping ────────────────────────────────────────────
STRATEGIC_AREAS = [
    ("commercial",    "Commercial & Business Development"),
    ("regulatory",    "Regulatory & Compliance"),
    ("legal",         "Legal Proceedings"),
    ("financial",     "Financial & Capital Management"),
    ("jv_partner",    "JV & Partner Relations"),
    ("operational",   "Operational & Technical"),
    ("stakeholder",   "Stakeholder & Investor Relations"),
    ("other",         "Other Matters"),
]
AREA_ORDER = {a[0]: i for i, a in enumerate(STRATEGIC_AREAS)}
AREA_LABELS = {a[0]: a[1] for a in STRATEGIC_AREAS}


# ─────────────────────────────────────────────
# AUTH  (identical to board_briefing.py)
# ─────────────────────────────────────────────

def _load_cache():
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text())
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize())
        try:
            CACHE_FILE.chmod(0o600)
        except Exception:
            pass

def _build_app(cache):
    if not CLIENT_ID:
        raise RuntimeError("OUTLOOK_CLIENT_ID not set. Add it to your .env file.")
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

def get_access_token() -> str:
    cache    = _load_cache()
    app      = _build_app(cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    raise RuntimeError("Token missing or expired.\nRun:  py matter_log.py setup")

def setup_auth():
    if not CLIENT_ID:
        print("\n❌  OUTLOOK_CLIENT_ID not found in .env")
        return
    cache    = _load_cache()
    app      = _build_app(cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            print(f"\n✅  Already authenticated as: {accounts[0]['username']}\n")
            return
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"\n❌  Failed: {json.dumps(flow, indent=2)}\n")
        return
    print(f"\n  1. Open:  https://microsoft.com/devicelogin")
    print(f"  2. Enter: {flow['user_code']}")
    print(f"  3. Sign in with your Microsoft account\n")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _save_cache(cache)
        print("✅  Authenticated successfully!\n")
    else:
        print(f"\n❌  {result.get('error_description', result)}\n")


# ─────────────────────────────────────────────
# API RETRY
# ─────────────────────────────────────────────

def _api_call_with_retry(client, max_retries: int = 3,
                         retry_delay: float = 12.0, **kwargs):
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return client.messages.create(**kwargs)
        except Exception as e:
            last_exc = e
            err_str  = str(e).lower()
            retriable = any(k in err_str for k in (
                "connection", "timeout", "rate", "429", "500", "502", "503", "overloaded"
            ))
            if not retriable or attempt == max_retries:
                raise
            is_rate = any(k in err_str for k in ("rate", "429"))
            wait    = (60 if is_rate else retry_delay) * attempt
            print(f" ⟳ retry {attempt}/{max_retries - 1} in {wait:.0f}s", end=" ", flush=True)
            time.sleep(wait)
    raise last_exc


# ─────────────────────────────────────────────
# ROBUST JSON PARSER
# ─────────────────────────────────────────────

def _robust_json_parse(text: str):
    # Strip markdown code fences (```json ... ```) before any parse attempt
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else text[3:]
    if text.endswith("```"):
        text = text[:-3].rstrip()
    text = text.strip()

    text = text.strip()
    for pat, repl in [(r'^```(?:json)?\s*', ''), (r'\s*```$', '')]:
        text = re.sub(pat, repl, text, flags=re.MULTILINE)
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    for sc, ec in [('{', '}'), ('[', ']')]:
        start = text.find(sc)
        end   = text.rfind(ec)
        if start != -1 and end > start:
            try:
                return json.loads(text[start:end + 1])
            except json.JSONDecodeError:
                pass

    # Strategy 3: truncated recovery for objects AND arrays
    if text.startswith('{') or text.startswith('['):
        is_array = text.startswith('[')
        last_safe = -1
        depth = 0
        in_string = False
        escape_next = False
        for i, c in enumerate(text):
            if escape_next:
                escape_next = False
                continue
            if c == '\\':
                escape_next = True
                continue
            if c == '"' and not escape_next:
                in_string = not in_string
                continue
            if in_string:
                continue
            if c in '{[':
                depth += 1
            elif c in '}]':
                depth -= 1
                if is_array and depth == 0:
                    # Found the complete array close
                    try:
                        return json.loads(text[:i + 1])
                    except json.JSONDecodeError:
                        pass
                elif is_array and depth == 1:
                    # Found the end of a complete object inside the array
                    last_safe = i + 1
                elif not is_array and depth == 1:
                    last_safe = i + 1
        # If truncated, close the array/object with what we have
        if last_safe > 0:
            truncated = text[:last_safe].rstrip().rstrip(',')
            closer = '\n]' if is_array else '\n}'
            try:
                return json.loads(truncated + closer)
            except json.JSONDecodeError:
                pass

    raise json.JSONDecodeError("All repair strategies failed", text, 0)


# ─────────────────────────────────────────────
# TOKEN BUDGET
# ─────────────────────────────────────────────

def _estimate_tokens(text: str) -> int:
    return int(len(text) / _CHARS_PER_TOKEN)

def _split_oversized_batch(emails: list, prompt_overhead: int) -> list:
    budget      = _MAX_BATCH_TOKENS - prompt_overhead
    sub_batches = []
    current     = []
    current_tok = 0
    for email in emails:
        body        = email.get('body', '')
        email_text  = (
            f"Subject: {email.get('subject','')}\n"
            f"From: {email.get('from_name','')}\n"
            f"Body: {body[:MAX_BODY_CHARS]}\n"
        )
        email_tokens = _estimate_tokens(email_text)
        if current and current_tok + email_tokens > budget:
            sub_batches.append(current)
            current     = []
            current_tok = 0
        current.append(email)
        current_tok += email_tokens
    if current:
        sub_batches.append(current)
    return sub_batches


# ─────────────────────────────────────────────
# FOLDER DISCOVERY
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

def get_all_folders(token: str) -> list:
    queue, results, seen = [], [], set()

    def enqueue(fid, fpath):
        if fid not in seen:
            seen.add(fid)
            queue.append((fid, fpath))
            results.append((fid, fpath))

    enqueue("inbox",     "Inbox")
    enqueue("sentitems", "Sent Items")

    print("     Discovering folders…")
    url    = f"{GRAPH_BASE}/me/mailFolders"
    params = {"$top": 100, "$select": "id,displayName"}
    while url:
        try:
            resp = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                                params=params, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
            for f in data.get("value", []):
                fname = f.get("displayName", "")
                if fname.lower().replace(" ", "") not in SKIP_SYSTEM:
                    enqueue(f["id"], fname)
            url    = data.get("@odata.nextLink")
            params = {}
        except Exception as e:
            print(f"   ⚠️  Folder error: {e}")
            break

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

    print(f"     ✅  {len(results)} folders")
    return results


# ─────────────────────────────────────────────
# EMAIL FETCHING
# ─────────────────────────────────────────────

def _fetch_body(token: str, message_id: str) -> str:
    try:
        data         = requests.get(
            f"{GRAPH_BASE}/me/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "body"},
            timeout=REQUEST_TIMEOUT,
        ).json()
        content_type = data.get("body", {}).get("contentType", "text")
        body         = data.get("body", {}).get("content", "")
        if content_type == "html":
            body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.DOTALL)
            body = re.sub(r'<[^>]+>', ' ', body)
            body = re.sub(r'[ \t]+', ' ', body)
            body = re.sub(r'\n{3,}', '\n\n', body).strip()
        return body[:MAX_BODY_CHARS]
    except Exception as e:
        return f"[body unavailable: {e}]"

def _parse_msg(msg: dict, folder_path: str) -> dict:
    sender  = msg.get("from", {}).get("emailAddress", {})
    to_list = [
        r.get("emailAddress", {}).get("name", "") or r.get("emailAddress", {}).get("address", "")
        for r in msg.get("toRecipients", [])[:5]
    ]
    ts = msg.get("receivedDateTime") or msg.get("sentDateTime", "")
    return {
        "id":         msg.get("id", ""),
        "subject":    msg.get("subject", "(no subject)").strip(),
        "from_name":  sender.get("name", ""),
        "from_email": sender.get("address", ""),
        "to":         ", ".join(to_list),
        "datetime":   ts[:19].replace("T", " ") if ts else "",
        "preview":    msg.get("bodyPreview", "")[:300].strip(),
        "importance": msg.get("importance", "normal"),
        "has_attach": msg.get("hasAttachments", False),
        "folder":     folder_path,
        "body":       "",
    }

def fetch_folder_emails(token: str, folder_id: str, folder_path: str,
                        since: datetime.datetime) -> list:
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")
    results   = []
    for date_field in ("receivedDateTime", "sentDateTime"):
        url    = f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
        params = {
            "$filter":  f"{date_field} ge {since_iso}",
            "$select":  "id,subject,from,toRecipients,receivedDateTime,sentDateTime,"
                        "hasAttachments,importance,isRead,bodyPreview",
            "$orderby": f"{date_field} desc",
            "$top":     MAX_EMAILS_PER_FOLDER,
        }
        try:
            resp = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                                params=params, timeout=REQUEST_TIMEOUT)
            if resp.status_code in (400, 404, 501):
                continue
            resp.raise_for_status()
            for msg in resp.json().get("value", []):
                mid = msg.get("id", "")
                if mid and not any(e["id"] == mid for e in results):
                    results.append(_parse_msg(msg, folder_path))
            if results:
                break
        except Exception:
            continue
    return results

def run_scan(token: str, lookback_days: int) -> list:
    since    = datetime.datetime.utcnow() - datetime.timedelta(days=lookback_days)
    since_dt = since.replace(hour=0, minute=0, second=0, microsecond=0)

    print(f"\n  📁  Folder discovery")
    folders = get_all_folders(token)
    print(f"       {len(folders)} folders to scan\n")

    print("  📬  Fetching emails…")
    all_msgs = {}
    for folder_id, folder_path in folders:
        msgs = fetch_folder_emails(token, folder_id, folder_path, since_dt)
        new  = 0
        for m in msgs:
            if m["id"] not in all_msgs:
                all_msgs[m["id"]] = m
                new += 1
        if new:
            print(f"     [{folder_path}]  {new} emails")
        if len(all_msgs) >= MAX_EMAILS_TOTAL:
            print(f"\n  ⚠️  Cap of {MAX_EMAILS_TOTAL} reached — stopping")
            break

    emails = sorted(all_msgs.values(), key=lambda e: e["datetime"], reverse=True)
    print(f"\n  ✅  {len(emails)} unique emails found\n")

    if emails:
        print("  📄  Fetching full bodies…")
        for i, email in enumerate(emails):
            print(f"       {i+1}/{len(emails)}: {email['subject'][:60]}", end=" ", flush=True)
            email["body"] = _fetch_body(token, email["id"])
            print("✓")
        print()

    return emails


# ─────────────────────────────────────────────
# EMAIL CONTEXT BUILDER
# ─────────────────────────────────────────────

def build_email_context(emails: list) -> str:
    lines = [f"TOTAL EMAILS: {len(emails)}", ""]
    for i, e in enumerate(emails):
        lines += [
            f"{'='*60}", f"EMAIL {i+1}",
            f"Date:    {e['datetime']}",
            f"From:    {e['from_name']} <{e['from_email']}>",
            f"To:      {e['to']}",
            f"Subject: {e['subject']}",
            f"Folder:  {e['folder']}",
        ]
        if e.get("has_attach"):
            lines.append("Attachments: YES")
        if e.get("importance") == "high":
            lines.append("Importance: HIGH")
        lines += [f"\nBODY:\n{e.get('body', '[no body]')}", ""]
    return "\n".join(lines)


# ─────────────────────────────────────────────
# BATCH EXTRACTION — more granular than board_briefing.py
# ─────────────────────────────────────────────

EXTRACT_PROMPT = """You are a senior executive assistant to Doug McAlpine, Managing Director of
{company_name} ({ticker}).

{context_section}{briefing_section}TODAY: {date_to}
PERIOD: {date_from} to {date_to}
EMAIL BATCH: {batch_num} of {total_batches}

YOUR TASK:
Extract every board-relevant topic from the emails below. This is a FACT-GATHERING
pass only — do NOT categorise, do NOT assess significance, do NOT write narratives.
Just capture the raw facts as concisely as possible.

WHAT TO CAPTURE:
- Commercial negotiations, counterparty discussions, deal terms
- Regulatory notices, compliance deadlines, government interactions
- Legal proceedings, court filings, lawyer correspondence
- Financial matters: cash calls, invoices, tax, R&D claims, capital raises
- Project milestones, engineering studies, operational updates
- Stakeholder and investor interactions

WHAT TO EXCLUDE:
- Routine scheduling without substantive outcome
- Travel/accommodation bookings
- Newsletter subscriptions and automated notifications
- Invoice reminders under $5,000 unless part of a pattern
- Generic IT/system notifications

CRITICAL: Capture ALL specific figures (dollar amounts, share counts, dates,
volumes, percentages). Do not round or omit numbers.

RESPOND ONLY WITH VALID JSON — a single array, no markdown fences:

[
  {{
    "topic":           "Short descriptive title — specific enough to identify the matter",
    "facts":           "2-4 sentences: what happened, what was decided, key positions of each party. Just facts, no opinion.",
    "key_figures":     ["$3.6M claim", "989,351 shares", "30 June 2026 deadline"],
    "parties":         ["Person or organisation name"],
    "date_range":      "e.g. 1-7 May 2026",
    "source_subjects": ["Email subject line 1", "Email subject line 2"]
  }}
]

Return [] only if this batch genuinely contains zero board-relevant content.

────────────────────────────────────────────────────────────
EMAILS ({email_count} in this batch):
────────────────────────────────────────────────────────────

{email_ctx}
"""
def extract_matters_batch(client, emails: list, batch_num: int,
                           total_batches: int, context_block: str,
                           briefing_text: str, lookback_days: int) -> list:
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today() - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")
    email_ctx = build_email_context(emails)

    # Escape curly braces in dynamic content to prevent .format() KeyErrors
    # Email bodies and context blocks frequently contain { and } characters
    # (JSON payloads, code snippets, template markers) which .format()
    # misinterprets as field names, silently failing the extraction.
    def _esc(s):
        return s.replace("{", "{{").replace("}", "}}")

    prompt = EXTRACT_PROMPT.format(
        company_name     = COMPANY_NAME,
        ticker           = COMPANY_TICKER,
        context_section  = _esc(f"{context_block}\n") if context_block else "",
        briefing_section = _esc(f"STANDING INSTRUCTIONS:\n{briefing_text}\n\n") if briefing_text else "",
        date_to          = date_to,
        date_from        = date_from,
        batch_num        = batch_num,
        total_batches    = total_batches,
        email_count      = len(emails),
        email_ctx        = _esc(email_ctx),
    )

    try:
        msg = _api_call_with_retry(
            client,
            model      = "claude-sonnet-4-6",
            max_tokens = 16000,
            messages   = [{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # Strip markdown code fences (```json ... ```) that Claude sometimes adds
        if raw.startswith("```"):
            raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
        if raw.endswith("```"):
            raw = raw[:-3].rstrip()
        try:
            matters = _robust_json_parse(raw)
        except json.JSONDecodeError as e:
            print(f"⚠️  parse error: {e}")
            print(f"       Response starts with: {repr(raw[:300])}")
            print(f"       Response length: {len(raw)} chars")
            return []
        if not isinstance(matters, list):
            matters = []
        n = len(matters)
        print(f"✅  {n} matter{'s' if n != 1 else ''}")
        return matters
    except Exception as e:
        print(f"⚠️  failed: {e}")
        return []


# ─────────────────────────────────────────────
# MATTER DEDUPLICATION AND MERGING
# ─────────────────────────────────────────────


# ─────────────────────────────────────────────
# STEP 2: CONSOLIDATE — group related items into distinct matters
# ─────────────────────────────────────────────

CONSOLIDATE_PROMPT = """You are helping consolidate a raw list of {n_items} email extractions from
{company_name} ({ticker}) into distinct board-level matters.

Many of these items relate to the SAME underlying matter — for example:
- "Landholder CCA — Affidavit Exchange", "Land Court — Compensation Statement",
  and "Law Firm Invoice — Landholder Dispute" are all ONE matter: the landholder litigation.
- "JV Partner Cash Call", "JV Partner JOA Rights Reserved", and "JV Partner Payment Plan"
  are all ONE matter: the JV dispute.

YOUR TASK:
1. Group related items into distinct matters. Be aggressive about consolidating —
   if two items involve the same counterparty and the same underlying issue, they
   are ONE matter regardless of whether the emails discussed legal, financial, or
   operational aspects.

2. For each consolidated matter, merge ALL facts from the constituent items into
   a single comprehensive factual summary.

3. Assess each consolidated matter:
   - area: commercial | regulatory | legal | financial | jv_partner | operational | stakeholder | other
     (choose the PRIMARY area — don't split one matter across categories)
   - significance: high | medium | low
     HIGH = requires board decision, material financial exposure (>$50k), or time-critical deadline
     MEDIUM = board should be aware, active management required
     LOW = routine, noting only
   - requires_decision: true if the board needs to make a decision or provide guidance

RESPOND ONLY WITH VALID JSON — a single array, no markdown fences:

[
  {{
    "matter":            "Clear, specific matter title",
    "area":              "commercial | regulatory | legal | financial | jv_partner | operational | stakeholder | other",
    "significance":      "high | medium | low",
    "merged_facts":      "Comprehensive factual summary merging all related items — 4-8 sentences covering the full picture. Include ALL figures, dates, and party positions.",
    "key_figures":       ["All dollar amounts, share counts, dates, volumes from constituent items"],
    "key_parties":       ["All parties involved"],
    "date_range":        "Earliest to latest date across all constituent items",
    "next_steps":        "What happens next in THIRD PERSON — State Gas will / Management will / The board is asked to...",
    "requires_decision": true or false,
    "decision_detail":   "If true: what exactly needs to be decided",
    "source_subjects":   ["Up to 5 most relevant email subject lines from across all constituent items"],
    "constituent_count": 3
  }}
]

RAW EXTRACTED ITEMS:
{items_json}
"""


def consolidate_matters(client, all_items: list) -> list:
    """
    Step 2: Send ALL extracted items to Claude for grouping and deduplication.
    If the list is too large, process in two passes.
    """
    if not all_items:
        return []

    def _esc(s):
        return s.replace("{", "{{").replace("}", "}}")

    # Compact the items for the prompt (remove verbose fields)
    compact = []
    for item in all_items:
        compact.append({
            "topic":       item.get("topic", item.get("matter", "")),
            "facts":       item.get("facts", item.get("narrative", ""))[:500],
            "key_figures": item.get("key_figures", []),
            "parties":     item.get("parties", item.get("key_parties", [])),
            "date_range":  item.get("date_range", ""),
            "source_subjects": item.get("source_subjects", [])[:2],
        })

    items_json = json.dumps(compact, indent=1, ensure_ascii=False)

    # Check if we need to split (rough estimate: >120k chars = split)
    if len(items_json) > 120000:
        mid = len(compact) // 2
        print(f"       Large item set ({len(compact)} items) — consolidating in 2 passes…")
        part1 = consolidate_matters(client, all_items[:mid])
        time.sleep(30)
        part2 = consolidate_matters(client, all_items[mid:])
        # Final merge pass
        return consolidate_matters(client, part1 + part2)

    prompt = CONSOLIDATE_PROMPT.format(
        company_name = COMPANY_NAME,
        ticker       = COMPANY_TICKER,
        n_items      = len(compact),
        items_json   = _esc(items_json),
    )

    print(f"  🔗  Consolidating {len(all_items)} extracted items into distinct matters…", end=" ", flush=True)
    try:
        msg = _api_call_with_retry(
            client,
            model      = "claude-sonnet-4-6",
            max_tokens = 16000,
            messages   = [{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # Strip markdown code fences
        if raw.startswith("```"):
            raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
        if raw.endswith("```"):
            raw = raw[:-3].rstrip()
        matters = _robust_json_parse(raw)
        if not isinstance(matters, list):
            matters = []
        print(f"✅  {len(matters)} distinct matters")
        return matters
    except json.JSONDecodeError as e:
        print(f"⚠️  consolidation parse error: {e}")
        print(f"       Response starts with: {repr(raw[:300])}")
        # Fall back to the raw items without consolidation
        return all_items
    except Exception as e:
        print(f"⚠️  consolidation failed: {e} — using raw items")
        return all_items

ENRICH_PROMPT = """You are helping Doug McAlpine, Managing Director of {company_name} ({ticker}),
finalise a working matter log covering {n_emails} emails from {date_from} to {date_to}.

Below are {n_matters} consolidated matters. For each matter, your task is to:

1. REWRITE the narrative in Doug's voice — first person, confident, conversational but
   professional. "I have been engaging with..." not "Management has been engaging with..."
   Target 4-6 sentences per matter. Preserve all specific figures, names, and dates.

2. SHARPEN the next_steps — make it concrete but write in THIRD PERSON.
   "Management will..." or "State Gas intends to..." or "The board is asked to..."
   NOT first person. Next steps appear in a formal board paper column.

3. ADD a "doug_note" field — one sentence of candid personal assessment:
   what Doug is most concerned about, most optimistic about, or most uncertain about
   regarding this matter. Base this only on what the facts reveal.

DO NOT change the matter title, area, significance, or requires_decision fields.
These have already been assessed.

RESPOND ONLY WITH VALID JSON — an array with one object per matter:

[
  {{
    "matter":        "exact matter title from input",
    "narrative":     "rewritten 4-6 sentence first-person narrative",
    "next_steps":    "rewritten concrete first-person next steps",
    "doug_note":     "one candid sentence about Doug's personal assessment"
  }}
]

Preserve the exact matter titles from the input so they can be matched back.

INPUT MATTERS:
{matters_json}
"""
def enrich_narratives(client, matters: list, emails: list, lookback_days: int) -> list:
    """
    Second Claude pass: rewrites all narratives in Doug's voice and adds
    personal notes. Processes in chunks of 15 matters to stay within limits.
    """
    if not matters:
        return matters

    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today() - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    # Process in chunks of 15 matters
    CHUNK_SIZE    = 15
    enriched_map  = {}
    chunks        = [matters[i:i+CHUNK_SIZE] for i in range(0, len(matters), CHUNK_SIZE)]

    print(f"\n  ✍️   Enriching narratives across {len(chunks)} chunk(s)…")

    for ci, chunk in enumerate(chunks, 1):
        print(f"       Chunk {ci}/{len(chunks)}: {len(chunk)} matters … ", end="", flush=True)

        # Compact input — just title, narrative, next_steps, key_figures
        compact = [
            {
                "matter":     m.get("matter", ""),
                "area":       m.get("area", ""),
                "narrative":  (m.get("merged_facts") or m.get("narrative") or "")[:1500],
                "next_steps": m.get("next_steps", ""),
                "key_figures": m.get("key_figures", []),
                "key_parties": m.get("key_parties", []),
                "requires_decision": m.get("requires_decision", False),
                "decision_detail":   m.get("decision_detail", ""),
            }
            for m in chunk
        ]

        # Escape curly braces in matters JSON (contains user content with { } chars)
        matters_json_safe = json.dumps(compact, indent=2, ensure_ascii=False).replace("{", "{{").replace("}", "}}")
        prompt = ENRICH_PROMPT.format(
            company_name  = COMPANY_NAME,
            ticker        = COMPANY_TICKER,
            n_emails      = len(emails),
            date_from     = date_from,
            date_to       = date_to,
            n_matters     = len(chunk),
            matters_json  = matters_json_safe,
        )

        try:
            msg = _api_call_with_retry(
                client,
                model      = "claude-sonnet-4-6",
                max_tokens = 6000,
                messages   = [{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            # Strip markdown code fences
            if raw.startswith("```"):
                raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
            if raw.endswith("```"):
                raw = raw[:-3].rstrip()
            try:
                enriched = _robust_json_parse(raw)
            except json.JSONDecodeError as e:
                print(f"⚠️  parse error: {e} — keeping originals for this chunk")
                enriched = []

            if isinstance(enriched, list):
                for e_item in enriched:
                    title = e_item.get("matter", "").strip()
                    if title:
                        enriched_map[title.lower()] = e_item
            print(f"✅")

        except Exception as e:
            print(f"⚠️  {e} — keeping originals for this chunk")

        # Pause between enrichment chunks
        if ci < len(chunks):
            time.sleep(8)

    # Apply enrichments back to matters
    result = []
    for m in matters:
        key   = m.get("matter", "").strip().lower()
        enr   = enriched_map.get(key, {})
        if enr:
            m["narrative"]   = enr.get("narrative",  m.get("narrative", m.get("merged_facts", "")))
            m["next_steps"]  = enr.get("next_steps", m.get("next_steps", ""))
            m["doug_note"]   = enr.get("doug_note",  "")
            m["merge_with"]  = enr.get("merge_with", None)
        # Ensure narrative is populated even if enrichment didn't cover this matter
        if not m.get("narrative") and m.get("merged_facts"):
            m["narrative"] = m["merged_facts"]
        result.append(m)

    return result


# ─────────────────────────────────────────────
# MARKDOWN WRITER
# ─────────────────────────────────────────────

SIG_EMOJI  = {"high": "🔴", "medium": "🟡", "low": "🟢"}
SIG_LABEL  = {"high": "High", "medium": "Medium", "low": "Low"}

def generate_matter_log(matters: list, emails: list,
                        lookback_days: int, run_date: str) -> str:
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today() - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")
    now       = datetime.datetime.now().strftime("%d %B %Y, %H:%M")

    high   = sum(1 for m in matters if m.get("significance") == "high")
    medium = sum(1 for m in matters if m.get("significance") == "medium")
    low    = sum(1 for m in matters if m.get("significance") == "low")
    decs   = sum(1 for m in matters if m.get("requires_decision"))

    lines = [
        f"# State Gas Limited — Matter Log",
        f"",
        f"> **Period:** {date_from} → {date_to}  ",
        f"> **Emails reviewed:** {len(emails)}  ",
        f"> **Matters identified:** {len(matters)} total "
        f"({high} high · {medium} medium · {low} low)  ",
        f"> **Requiring decision:** {decs}  ",
        f"> **Generated:** {now}  ",
        f"> **Status:** Working document — not for distribution",
        f"",
        f"---",
        f"",
    ]

    # Decisions required summary at top
    decision_items = [m for m in matters if m.get("requires_decision")]
    if decision_items:
        lines += [
            f"## ⚡ Decisions Required ({len(decision_items)})",
            f"",
            f"_These matters require a decision or formal guidance before or at the next board meeting._",
            f"",
        ]
        for m in decision_items:
            lines += [
                f"**{m.get('matter', '')}**  ",
                f"{m.get('decision_detail', m.get('next_steps', ''))}",
                f"",
            ]
        lines += ["---", ""]

    # Group by area
    by_area: dict = {}
    for m in matters:
        area = m.get("area", "other")
        by_area.setdefault(area, []).append(m)

    # Render each area in defined order
    for area_id, area_label in STRATEGIC_AREAS:
        area_matters = by_area.get(area_id, [])
        if not area_matters:
            continue

        h_count = sum(1 for m in area_matters if m.get("significance") == "high")
        badge   = f"  ·  🔴 {h_count} high" if h_count else ""

        lines += [
            f"## {area_label}{badge}",
            f"",
        ]

        for m in area_matters:
            sig       = m.get("significance", "low")
            sig_emoji = SIG_EMOJI.get(sig, "⚪")
            sig_label = SIG_LABEL.get(sig, "")
            matter    = m.get("matter", "")
            date_rng  = m.get("date_range", "")
            parties   = m.get("key_parties", [])
            figures   = m.get("key_figures", [])
            narrative = m.get("narrative", "")
            next_steps = m.get("next_steps", "")
            doug_note  = m.get("doug_note", "")
            sources    = m.get("source_subjects", [])
            dec_req    = m.get("requires_decision", False)

            # Matter heading
            dec_flag = "  ⚡ _Decision required_" if dec_req else ""
            lines += [f"### {sig_emoji} {matter}{dec_flag}", ""]

            # Metadata line
            meta_parts = []
            if date_rng:
                meta_parts.append(f"_{date_rng}_")
            if parties:
                meta_parts.append(f"**Parties:** {', '.join(parties[:6])}")
            if meta_parts:
                lines += ["  ".join(meta_parts), ""]

            # Key figures callout
            if figures:
                lines += [
                    f"**Key figures:** {' · '.join(str(f) for f in figures[:8])}",
                    "",
                ]

            # Narrative
            if narrative:
                lines += [narrative, ""]

            # Decision detail
            if dec_req and m.get("decision_detail"):
                lines += [
                    f"> **Decision required:** {m['decision_detail']}",
                    "",
                ]

            # Next steps
            if next_steps:
                lines += [f"**Next steps:** {next_steps}", ""]

            # Doug's candid note (formatted distinctively)
            if doug_note:
                lines += [
                    f"<details>",
                    f"<summary>💭 Personal note</summary>",
                    f"",
                    f"_{doug_note}_",
                    f"",
                    f"</details>",
                    "",
                ]

            # Source emails
            if sources:
                lines += [
                    f"<details>",
                    f"<summary>Source emails</summary>",
                    f"",
                ]
                for s in sources[:5]:
                    lines.append(f"- {s}")
                lines += ["", "</details>", ""]

            lines.append("---")
            lines.append("")

    # Footer
    lines += [
        "",
        f"_Matter Log generated by matter_log.py — {len(emails)} emails reviewed over {lookback_days} days._",
        f"_This document is a working tool for the MD only and is not suitable for distribution._",
    ]

    return "\n".join(lines)


# ─────────────────────────────────────────────
# MAIN RUN
# ─────────────────────────────────────────────

def _find_node() -> str | None:
    """
    Locate the node executable on Windows and Unix.
    Tries shutil.which first, then common Windows install paths.
    Returns the full path string, or None if not found.
    """
    import shutil
    node = shutil.which("node") or shutil.which("node.exe")
    if node:
        return node
    # Common Windows install locations
    import os
    candidates = [
        r"C:\Program Files\nodejs\node.exe",
        r"C:\Program Files (x86)\nodejs\node.exe",
        os.path.expandvars(r"%APPDATA%\npm\node.exe"),
        os.path.expandvars(r"%ProgramFiles%\nodejs\node.exe"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _run_node(js: Path, md_path: Path, label: str) -> bool:
    """
    Run a node.js generator script, handling Windows PATH issues gracefully.
    Returns True on success, False on failure.
    Prints full stdout+stderr on failure so errors are always visible.
    """
    import subprocess

    if not js.exists():
        print(f"\n  ⚠️  {js.name} not found in {js.parent}")
        print(f"       Copy {js.name} into the same folder as matter_log.py")
        return False

    node = _find_node()
    if not node:
        print(f"\n  ⚠️  node.js not found — skipping {label}")
        print(f"       Install node.js from https://nodejs.org")
        print(f"       Then open a new terminal so the PATH is refreshed")
        return False

    # Try with explicit node path first (avoids Windows PATH lookup)
    for cmd, use_shell in [([node, str(js), "--data", str(md_path)], False),
                            (f'node "{js}" --data "{md_path}"', True)]:
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True,
                shell=use_shell,
                cwd=str(js.parent)   # run from the JS file directory so require() finds node_modules
            )
            if result.returncode == 0:
                return True
            # Print full error output — never truncate
            out = (result.stdout or "").strip()
            err = (result.stderr or "").strip()
            print(f"\n  ⚠️  {label} failed (exit {result.returncode})")
            if err:
                print(f"       stderr: {err}")
            if out:
                print(f"       stdout: {out}")
            # If this was the shell=False attempt and it looks like a PATH issue, try shell=True
            if not use_shell and "Cannot find module" not in (err + out):
                return False
            return False
        except FileNotFoundError:
            if use_shell:
                print(f"\n  ⚠️  node.js not accessible — skipping {label}")
                return False
            continue  # try shell=True fallback
        except Exception as e:
            print(f"\n  ⚠️  {label} error: {e}")
            return False

    return False


def _generate_pptx(md_path: Path):
    """Run the PowerPoint generator against the given markdown file."""
    js = _here / "generate_matter_log_deck.js"
    print(f"  🖥️   Generating PowerPoint…", end=" ", flush=True)
    if _run_node(js, md_path, "PowerPoint"):
        print("✅")
    else:
        print("skipped")


def _generate_docx(matters: list, emails: list, lookback_days: int, out_path: Path):
    """
    Generate a landscape A4 Word document directly from the matters list.
    Uses python-docx — no Node.js dependency.
    Install: pip install python-docx
    """
    try:
        from docx import Document
        from docx.shared import Inches, Pt, Cm, RGBColor, Emu
        from docx.enum.section import WD_ORIENT
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn, nsdecls
        from docx.oxml import parse_xml
    except ImportError:
        print("  ⚠️  python-docx not installed — skipping Word output")
        print("       Install: pip install python-docx")
        return

    print(f"  📝  Generating Word document…", end=" ", flush=True)

    docx_path = out_path.with_suffix(".docx")
    date_to   = datetime.date.today().strftime("%d %B %Y")
    date_from = (datetime.date.today() - datetime.timedelta(days=lookback_days)).strftime("%d %B %Y")

    # Brand colours
    CRIMSON  = RGBColor(0xA5, 0x1C, 0x30)
    DARK     = RGBColor(0x1A, 0x1A, 0x1A)
    WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
    SLATE    = RGBColor(0x5A, 0x5A, 0x5A)
    GREEN    = RGBColor(0x1A, 0x6B, 0x38)
    AMBER    = RGBColor(0xB3, 0x5C, 0x00)
    RED      = RGBColor(0xA5, 0x1C, 0x30)
    ALT_BG   = "FAF5F5"
    SIG_CONF = {
        "high":   {"label": "HIGH", "bg": "A51C30", "fg": WHITE},
        "medium": {"label": "MED",  "bg": "B35C00", "fg": WHITE},
        "low":    {"label": "LOW",  "bg": "1A6B38", "fg": WHITE},
    }
    AREA_LABELS = {
        "commercial":"Commercial & BD", "regulatory":"Regulatory", "legal":"Legal",
        "financial":"Financial", "jv_partner":"JV / Partner", "operational":"Operational",
        "stakeholder":"Stakeholder & IR", "other":"Other",
    }
    AREA_ORDER = ["commercial","regulatory","legal","financial","jv_partner","operational","stakeholder","other"]

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_text(cell, text, bold=False, italic=False, size=9, color=DARK, alignment=None):
        cell.text = ""
        p = cell.paragraphs[0]
        if alignment:
            p.alignment = alignment
        run = p.add_run(str(text))
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.name = "Calibri"
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

    def trunc(s, n):
        if not s: return ""
        s = " ".join(str(s).split()).strip()
        return s if len(s) <= n else s[:n-1] + "…"

    # Sort matters: decisions first, then significance, then area
    sig_order  = {"high":0, "medium":1, "low":2}
    area_order = {a:i for i,a in enumerate(AREA_ORDER)}
    sorted_matters = sorted(matters, key=lambda m: (
        0 if m.get("requires_decision") else 1,
        sig_order.get(m.get("significance","low"), 2),
        area_order.get(m.get("area","other"), 99),
        m.get("matter",""),
    ))

    high_count = sum(1 for m in matters if m.get("significance") == "high")
    dec_count  = sum(1 for m in matters if m.get("requires_decision"))

    # Create document
    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    font  = style.font
    font.name = "Calibri"
    font.size = Pt(9)

    # Set landscape A4 for all sections
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width  = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin    = Cm(1.3)
    section.bottom_margin = Cm(1.3)
    section.left_margin   = Cm(1.3)
    section.right_margin  = Cm(1.3)

    # Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hr = hp.add_run("State Gas Limited — Matter Log")
    hr.bold = True
    hr.font.size = Pt(9)
    hr.font.color.rgb = DARK
    hr.font.name = "Calibri"
    hp.add_run("    ").font.size = Pt(9)
    hr2 = hp.add_run(f"{date_from} → {date_to}")
    hr2.font.size = Pt(8)
    hr2.font.color.rgb = SLATE
    hr2.font.name = "Calibri"
    # Crimson bottom border on header
    pPr = hp._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="A51C30"/></w:pBdr>')
    pPr.append(pBdr)

    # Footer
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fr = fp.add_run("WORKING DOCUMENT — NOT FOR DISTRIBUTION")
    fr.italic = True
    fr.font.size = Pt(7)
    fr.font.color.rgb = SLATE
    fr.font.name = "Calibri"

    # ── Cover page ─────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.space_before = Pt(36)
    r = p.add_run("STATE GAS LIMITED")
    r.bold = True; r.font.size = Pt(12); r.font.color.rgb = CRIMSON; r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run("Matter Log")
    r.bold = True; r.font.size = Pt(28); r.font.color.rgb = DARK; r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run(f"{date_from} → {date_to}")
    r.font.size = Pt(12); r.font.color.rgb = SLATE; r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run(f"{len(emails)} emails reviewed  ·  {len(matters)} matters identified  ·  {high_count} high significance  ·  {dec_count} decisions required")
    r.font.size = Pt(10); r.font.color.rgb = SLATE; r.font.name = "Calibri"

    p = doc.add_paragraph()
    p.space_before = Pt(24)
    r = p.add_run("WORKING DOCUMENT — NOT FOR DISTRIBUTION")
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = CRIMSON; r.font.name = "Calibri"

    doc.add_page_break()

    # ── Main table ─────────────────────────────────────────────────────
    # Column widths (approximate, in inches, total ~10.5" for landscape A4 content)
    col_widths = [Inches(0.45), Inches(1.1), Inches(1.8), Inches(4.7), Inches(1.8), Inches(0.55)]

    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.allow_autofit = False

    # Set column widths
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    # Header row
    headers = ["Sig", "Area", "Matter", "Update", "Next Steps", "Dec?"]
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_text(cell, h, bold=True, size=8, color=WHITE)
        set_cell_shading(cell, "1A1A1A")
        cell.width = col_widths[i]

    # Data rows grouped by area
    last_area = None
    row_idx   = 0

    for m in sorted_matters:
        area = m.get("area", "other")

        # Area break row
        if area != last_area:
            area_row = table.add_row()
            # Merge all cells for area header
            area_row.cells[0].merge(area_row.cells[5])
            merged_cell = area_row.cells[0]
            set_cell_text(merged_cell, (AREA_LABELS.get(area, area)).upper(), bold=True, size=9, color=WHITE)
            set_cell_shading(merged_cell, "A51C30")
            last_area = area

        # Data row
        row = table.add_row()
        sig = m.get("significance", "medium")
        sc  = SIG_CONF.get(sig, SIG_CONF["medium"])
        dec_req = m.get("requires_decision", False)
        row_bg  = "FFFFFF" if row_idx % 2 == 0 else ALT_BG

        # Build narrative text
        narr_parts = [m.get("narrative", "")]
        if m.get("key_figures"):
            kf = m["key_figures"]
            if isinstance(kf, list):
                kf = " · ".join(str(f) for f in kf)
            narr_parts.append(f"[{kf}]")
        narrative = trunc(" ".join(p for p in narr_parts if p), 1200)

        # Build next steps text
        next_parts = [m.get("next_steps", "")]
        if dec_req and m.get("decision_detail"):
            next_parts.append(f"Decision required: {m['decision_detail']}")
        next_text = trunc(" ".join(p for p in next_parts if p), 400)

        # Sig cell
        set_cell_text(row.cells[0], sc["label"], bold=True, size=8, color=sc["fg"],
                      alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(row.cells[0], sc["bg"])

        # Area cell
        set_cell_text(row.cells[1], AREA_LABELS.get(area, area), italic=True, size=8, color=SLATE)
        set_cell_shading(row.cells[1], row_bg)

        # Matter cell — title + date + parties
        cell_m = row.cells[2]
        cell_m.text = ""
        p = cell_m.paragraphs[0]
        matter_color = CRIMSON if sig == "high" else DARK
        r = p.add_run(m.get("matter", ""))
        r.bold = True; r.font.size = Pt(9); r.font.color.rgb = matter_color; r.font.name = "Calibri"
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(0)
        if m.get("date_range") or m.get("key_parties"):
            p2 = cell_m.add_paragraph()
            meta_parts = []
            if m.get("date_range"):
                meta_parts.append(str(m["date_range"]))
            if m.get("key_parties"):
                parties = m["key_parties"]
                if isinstance(parties, list):
                    parties = ", ".join(parties[:4])
                meta_parts.append(trunc(parties, 100))
            r2 = p2.add_run(" · ".join(meta_parts))
            r2.font.size = Pt(7); r2.font.color.rgb = SLATE; r2.font.name = "Calibri"; r2.italic = True
            p2.paragraph_format.space_before = Pt(1)
            p2.paragraph_format.space_after = Pt(1)
        set_cell_shading(cell_m, row_bg)

        # Narrative cell
        set_cell_text(row.cells[3], narrative, size=9, color=DARK)
        set_cell_shading(row.cells[3], row_bg)

        # Next steps cell
        set_cell_text(row.cells[4], next_text, size=9, bold=dec_req,
                      color=CRIMSON if dec_req else DARK)
        set_cell_shading(row.cells[4], row_bg)

        # Decision cell
        if dec_req:
            set_cell_text(row.cells[5], "YES", bold=True, size=8, color=WHITE,
                          alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(row.cells[5], "A51C30")
        else:
            set_cell_text(row.cells[5], "", size=8)
            set_cell_shading(row.cells[5], row_bg)

        # Set column widths on each row (python-docx requires this per-row)
        for i, width in enumerate(col_widths):
            row.cells[i].width = width

        row_idx += 1

    # Set thin borders on all cells
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            borders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>' +
                '<w:top w:val="single" w:sz="2" w:space="0" w:color="DDDDDD"/>' +
                '<w:bottom w:val="single" w:sz="2" w:space="0" w:color="DDDDDD"/>' +
                '<w:left w:val="single" w:sz="2" w:space="0" w:color="DDDDDD"/>' +
                '<w:right w:val="single" w:sz="2" w:space="0" w:color="DDDDDD"/>' +
                '</w:tcBorders>'
            )
            tcPr.append(borders)

    # Save
    doc.save(str(docx_path))
    size_kb = docx_path.stat().st_size // 1024
    print(f"✅  ({size_kb} KB)")


def cmd_run(lookback_days: int):
    print(f"\n{'='*60}")
    print(f"  STATE GAS — MATTER LOG GENERATOR")
    print(f"{'='*60}")
    print(f"  Lookback: {lookback_days} days")
    print(f"  Output:   StateGas_MatterLog_DATE.md")
    print(f"  Note:     This does NOT affect board_report_data.json")
    print(f"{'='*60}")

    try:
        token = get_access_token()

        # Load context if available
        ctx = BoardBriefingContext() if _CTX_AVAILABLE else None
        if ctx and not ctx.is_empty():
            runs = ctx.data.get("run_count", 0)
            print(f"\n  🧠  Context loaded (run #{runs})\n")
            context_block = ctx.build_prompt_block()
            briefing_text = ctx.load_context_briefing()
        else:
            context_block = ""
            briefing_text = ""
            print("\n  ℹ️  No prior context — run board_briefing.py build-context first\n")

        # Scan
        print(f"  Period: {(datetime.date.today() - datetime.timedelta(days=lookback_days)).strftime('%d %B %Y')} "
              f"→ {datetime.date.today().strftime('%d %B %Y')}\n")
        emails = run_scan(token, lookback_days)
        if not emails:
            print("  ⚠️  No emails found.")
            return

        # Batch extraction
        client  = anthropic.Anthropic()
        batches = [emails[i:i+MAX_EMAILS_PER_BATCH]
                   for i in range(0, len(emails), MAX_EMAILS_PER_BATCH)]

        # ── STEP 1: Extract facts from emails (no categorisation) ──
        print(f"  🔍  Step 1/3 — extracting facts across {len(batches)} batch(es)…\n")
        all_items = []
        for i, batch in enumerate(batches, 1):
            prompt_overhead = _estimate_tokens(context_block + briefing_text) + 2000
            sub_batches     = _split_oversized_batch(batch, prompt_overhead)
            if len(sub_batches) > 1:
                print(f"       Batch {i}: {len(batch)} emails → {len(sub_batches)} sub-batches")
            for si, sub in enumerate(sub_batches):
                label = f"{i}.{si+1}" if len(sub_batches) > 1 else str(i)
                print(f"       Batch {label}/{len(batches)}: {len(sub)} emails … ", end="", flush=True)
                items = extract_matters_batch(
                    client, sub, i, len(batches),
                    context_block, briefing_text, lookback_days
                )
                all_items.extend(items)

        print(f"\n  ✅  {len(all_items)} raw items extracted from {len(emails)} emails")

        # Pause before consolidation
        print(f"\n  ⏳  Waiting 30s before consolidation pass…")
        time.sleep(30)

        # ── STEP 2: Consolidate into distinct matters ──
        print(f"\n  🔗  Step 2/3 — consolidating into distinct matters…\n")
        consolidated = consolidate_matters(client, all_items)
        high = sum(1 for m in consolidated if m.get("significance") == "high")
        print(f"\n  ✅  {len(all_items)} items → {len(consolidated)} distinct matters ({high} high significance)")

        # Pause before enrichment
        print(f"\n  ⏳  Waiting 30s before enrichment pass…")
        time.sleep(30)

        # ── STEP 3: Enrich in Doug's voice ──
        print(f"\n  ✍️   Step 3/3 — enriching narratives in Doug's voice…")
        enriched = enrich_narratives(client, consolidated, emails, lookback_days)

        # Write Markdown output
        print(f"\n  📄  Writing outputs…", end=" ", flush=True)
        today    = datetime.date.today().strftime("%Y%m%d")
        out_path = OUTPUT_DIR / f"StateGas_MatterLog_{today}.md"
        content  = generate_matter_log(enriched, emails, lookback_days, today)
        out_path.write_text(content, encoding="utf-8")
        size_kb  = out_path.stat().st_size // 1024
        print(f"✅")

        # Generate PowerPoint and Word outputs
        _generate_pptx(out_path)
        _generate_docx(enriched, emails, lookback_days, out_path)

        decs = sum(1 for m in enriched if m.get("requires_decision"))
        print(f"\n{'='*60}")
        print(f"  COMPLETE")
        print(f"{'='*60}")
        docx_name = out_path.with_suffix(".docx").name
        print(f"  📋  Markdown:    {out_path.name}  ({size_kb} KB)")
        print(f"  📊  PowerPoint:  {out_path.with_suffix('.pptx').name}")
        print(f"  📝  Word doc:    {docx_name}")
        print(f"  Matters:     {len(enriched)} total, {high} high significance")
        if decs:
            print(f"  ⚡  Decisions:   {decs} items require board decision/guidance")
        print(f"\n  These files are for your review only.")
        print(f"  They do NOT affect board_report_data.json or the PowerPoint board deck.")
        print(f"{'='*60}\n")

    except RuntimeError as e:
        print(f"\n❌  {e}\n")
    except Exception as e:
        print(f"\n❌  Unexpected error: {e}\n")
        raise


# ─────────────────────────────────────────────
# TEST
# ─────────────────────────────────────────────

def cmd_test():
    print("\n🔍  Testing connection…\n")
    try:
        token = get_access_token()
        me    = requests.get(f"{GRAPH_BASE}/me",
                             headers={"Authorization": f"Bearer {token}"},
                             timeout=REQUEST_TIMEOUT).json()
        print(f"✅  Connected as: {me.get('displayName')} "
              f"({me.get('mail') or me.get('userPrincipalName')})\n")
        folders = get_all_folders(token)
        print(f"  📁  {len(folders)} folders discovered")
        if _CTX_AVAILABLE:
            ctx = BoardBriefingContext()
            print(f"  🧠  Context: run #{ctx.data.get('run_count',0)}, "
                  f"last updated {ctx.data.get('last_updated','never')}")
            print(f"  🗂️  Matters in context: {len(ctx.data.get('matters',{}))}")
            print(f"  👥  People in context:  {len(ctx.data.get('key_people',{}))}")
    except RuntimeError as e:
        print(f"❌  {e}\n")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="State Gas comprehensive matter log generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("command", nargs="?", default="help",
                        choices=["setup", "test", "run", "help"])
    parser.add_argument("--days", type=int, default=DEFAULT_LOOKBACK_DAYS,
                        help=f"Lookback window in days (default: {DEFAULT_LOOKBACK_DAYS})")
    args = parser.parse_args()

    if args.command == "setup":
        setup_auth()
    elif args.command == "test":
        cmd_test()
    elif args.command == "run":
        cmd_run(args.days)
    else:
        print(__doc__)
