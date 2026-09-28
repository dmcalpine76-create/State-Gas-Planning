"""
board_sources.py - evidence the board pack needs that is not in email.

  * The minutes of the last board meeting (decisions, requests, action items),
    from the board/minutes folder in the knowledge store (OneDrive). The MD
    saves them there; nothing else is searched.
  * Market and policy context for the period, from a web search.
  * The MD's running notes for the next meeting (board/notes_for_next_meeting.md).

Everything is best-effort: if a source can't be reached the pack still runs.
"""
import json
import datetime
from pathlib import Path

MINUTES_DIR = "board/minutes"      # in the knowledge store (OneDrive), filled by the MD


def _file_date(f: Path) -> str:
    """Meeting date from a leading YYYY-MM-DD in the file name, else the file's own date."""
    import re
    m = re.match(r"(\d{4}-\d{2}-\d{2})", f.name)
    if m:
        return m.group(1)
    return datetime.date.fromtimestamp(f.stat().st_mtime).isoformat()


def doc_text(f: Path) -> str:
    suf = f.suffix.lower()
    if suf == ".docx":
        from docx import Document
        d = Document(str(f))
        parts = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(parts)
    if suf in (".md", ".txt"):
        return f.read_text(encoding="utf-8", errors="ignore")
    if suf == ".pdf":
        try:
            from pypdf import PdfReader
            return "\n".join((pg.extract_text() or "") for pg in PdfReader(str(f)).pages)
        except Exception:
            return ""
    return ""


def last_minutes(store: Path, before: str) -> dict:
    """The most recent minutes in the store's board/minutes folder dated before the given ISO date."""
    d = Path(store) / MINUTES_DIR
    if not d.exists():
        return {}
    files = [f for f in d.iterdir() if f.is_file() and f.suffix.lower() in (".docx", ".md", ".txt", ".pdf")
             and _file_date(f) < before]
    if not files:
        return {}
    f = max(files, key=_file_date)
    return {"name": f.name, "folder": MINUTES_DIR, "saved": _file_date(f), "text": doc_text(f)[:16000]}


def market_context(client, since: str, until: str, model: str = None) -> str:
    """Policy and market developments in the window, via a web search."""
    import anthropic
    from core.llm import DEFAULT_MODEL
    prompt = f"""Find the developments between {since} and {until} that the board of an ASX-listed
Queensland coal seam gas explorer (Bowen Basin; developing a gas project and a heavy-duty
natural gas truck refuelling business) should have in mind. Cover: federal and Queensland gas
policy (including any domestic gas reservation scheme), east coast gas market and prices,
notable gas M&A or asset sales that act as value markers, and conditions for energy
micro-cap capital raising. Only include items published within that window - ignore anything
later. Give 4 to 8 items, each one line: date, what happened, why it matters to a Queensland
gas junior. Plain text, no preamble."""
    try:
        msg = client.messages.create(
            model=model or DEFAULT_MODEL, max_tokens=2000,
            tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 6}],
            messages=[{"role": "user", "content": prompt}])
        return "\n".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()[:6000]
    except Exception:
        return ""


def running_notes(store: Path) -> str:
    f = Path(store) / "board" / "notes_for_next_meeting.md"
    try:
        t = f.read_text(encoding="utf-8")
    except Exception:
        return ""
    body = "\n".join(l for l in t.splitlines() if l.strip() and not l.startswith("#")
                     and not l.lstrip().startswith(">"))
    return body.strip()[:8000]
