"""
board_sources.py - evidence the board pack needs that is not in email.

  * The minutes of the last board meeting (decisions, requests, action items)
    and the final board report, read from the company's board meeting folders
    in SharePoint. Where those live is set in the knowledge store
    (board/sources.json), not in this public code.
  * Market and policy context for the period, from a web search.
  * The MD's running notes for the next meeting (board/notes_for_next_meeting.md).

Everything is best-effort: if a source can't be reached the pack still runs.
"""
import io
import json
import datetime
import requests
from pathlib import Path
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"


def _cfg(store: Path) -> dict:
    f = Path(store) / "board" / "sources.json"
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return {}


class MeetingFolders:
    """The board meeting folders in SharePoint: Meetings/<year>/<month folder>/."""

    def __init__(self, token: str, cfg: dict):
        self.h = {"Authorization": f"Bearer {token}"}
        self.cfg = cfg
        r = requests.get(f"{GRAPH}/sites/{cfg['site_host']}:{cfg['site_path']}", headers=self.h, timeout=30)
        r.raise_for_status()
        self.site = r.json()["id"]

    def _children(self, rel: str) -> list:
        path = "/".join(quote(p) for p in rel.strip("/").split("/"))
        url = f"{GRAPH}/sites/{self.site}/drive/root:/{path}:/children"
        out, params = [], {"$select": "id,name,file,folder,lastModifiedDateTime,size", "$top": "200"}
        while url:
            r = requests.get(url, headers=self.h, params=params, timeout=60)
            if r.status_code == 404:
                return out
            r.raise_for_status()
            d = r.json()
            out += d.get("value", [])
            url, params = d.get("@odata.nextLink"), None
        return out

    def files(self, years) -> list:
        """Every file directly inside each meeting folder for the given years."""
        base = self.cfg.get("meetings_path", "").strip("/")
        out = []
        for y in years:
            for folder in self._children(f"{base}/{y}"):
                if "folder" not in folder:
                    continue
                for f in self._children(f"{base}/{y}/{folder['name']}"):
                    if "file" in f:
                        out.append({**f, "meeting_folder": f"{y}/{folder['name']}"})
        return out

    def download(self, item_id: str) -> bytes:
        r = requests.get(f"{GRAPH}/sites/{self.site}/drive/items/{item_id}/content",
                         headers=self.h, timeout=180)
        r.raise_for_status()
        return r.content


def docx_text(blob: bytes) -> str:
    from docx import Document
    d = Document(io.BytesIO(blob))
    parts = [p.text for p in d.paragraphs if p.text.strip()]
    for t in d.tables:
        for row in t.rows:
            parts.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(parts)


def last_minutes(token: str, store: Path, before: str) -> dict:
    """The most recent board minutes saved before the given ISO date."""
    cfg = _cfg(store)
    if not cfg.get("site_host"):
        return {}
    mf = MeetingFolders(token, cfg)
    y = int(before[:4])
    files = [f for f in mf.files([y - 1, y])
             if "minute" in f["name"].lower() and f["name"].lower().endswith(".docx")
             and f["lastModifiedDateTime"][:10] < before]
    if not files:
        return {}
    f = max(files, key=lambda x: x["lastModifiedDateTime"])
    return {"name": f["name"], "folder": f["meeting_folder"], "saved": f["lastModifiedDateTime"][:10],
            "text": docx_text(mf.download(f["id"]))[:16000]}


def final_decks(token: str, store: Path, after: str) -> list:
    """Board report decks saved to the meeting folders after the given date, newest first."""
    cfg = _cfg(store)
    if not cfg.get("site_host"):
        return []
    mf = MeetingFolders(token, cfg)
    y = int(after[:4])
    decks = [f for f in mf.files([y, y + 1])
             if f["name"].lower().endswith(".pptx") and "board report" in f["name"].lower()
             and f["lastModifiedDateTime"][:10] >= after]
    decks.sort(key=lambda x: x["lastModifiedDateTime"], reverse=True)
    return [{**d, "fetch": (lambda i=d["id"]: mf.download(i))} for d in decks]


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
