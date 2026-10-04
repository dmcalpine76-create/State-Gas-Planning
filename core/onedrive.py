"""
core/onedrive.py - reach the knowledge store in OneDrive from the cloud.

On the laptop the store is simply the synced folder. A cloud run has no
synced folder, so it copies the files it needs down into a temporary folder
through Microsoft Graph, runs exactly as it would on the laptop, and copies
back only the files it changed. Every upload checks the file has not changed
in OneDrive since it was downloaded (If-Match), so a laptop edit made in the
meantime is never overwritten: the upload stops with an error instead.
"""
import json
import requests
from pathlib import Path
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"
DEFAULT_PATH = "Documents/Current document editing/new AI projects/state gas knowledge"


class OneDriveStore:
    def __init__(self, token: str, drive_path: str, local_dir: Path):
        self.token = token
        self.base = drive_path.strip("/")
        self.local = Path(local_dir)
        self.local.mkdir(parents=True, exist_ok=True)
        self.etags = {}          # relative path -> eTag at download
        self.hashes = {}         # relative path -> content at download
        self.conflicts = []      # files left alone because OneDrive changed them mid-run

    def _url(self, rel: str, suffix: str = "") -> str:
        path = "/".join(quote(p) for p in f"{self.base}/{rel}".split("/"))
        return f"{GRAPH}/me/drive/root:/{path}{suffix}"

    def _h(self, extra=None):
        h = {"Authorization": f"Bearer {self.token}"}
        if extra:
            h.update(extra)
        return h

    def pull(self, rel: str, required: bool = False) -> bool:
        """Download one file into the local mirror. False if it doesn't exist."""
        meta = requests.get(self._url(rel), headers=self._h(), timeout=60)
        if meta.status_code == 404:
            if required:
                raise FileNotFoundError(f"store file missing: {rel}")
            return False
        meta.raise_for_status()
        info = meta.json()
        self.etags[rel] = info.get("eTag")
        data = requests.get(self._url(rel, ":/content"), headers=self._h(), timeout=120)
        data.raise_for_status()
        dest = self.local / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data.content)
        try:                    # keep OneDrive's modified date, so file dates mean the same in the cloud
            import os, datetime as _dt
            ts = _dt.datetime.fromisoformat(info["lastModifiedDateTime"].replace("Z", "+00:00")).timestamp()
            os.utime(dest, (ts, ts))
        except Exception:
            pass
        self.hashes[rel] = data.content
        return True

    def list(self, rel_dir: str) -> list:
        r = requests.get(self._url(rel_dir, ":/children"), headers=self._h(),
                         params={"$select": "name,file", "$top": "200"}, timeout=60)
        if r.status_code == 404:
            return []
        r.raise_for_status()
        return [c["name"] for c in r.json().get("value", []) if "file" in c]

    def push(self, rel: str, force_new: bool = False) -> bool:
        """Upload one local file if it changed. Refuses if OneDrive changed meanwhile."""
        src = self.local / rel
        if not src.exists():
            return False
        body = src.read_bytes()
        if not force_new and self.hashes.get(rel) == body:
            return False
        headers = {"Content-Type": "application/octet-stream"}
        if rel in self.etags and self.etags[rel]:
            headers["If-Match"] = self.etags[rel]
        r = requests.put(self._url(rel, ":/content"), headers=self._h(headers),
                         data=body, timeout=120)
        if r.status_code == 412:
            return self._resolve_conflict(rel, body)
        r.raise_for_status()
        self.etags[rel] = r.json().get("eTag")
        self.hashes[rel] = body
        return True

    def _resolve_conflict(self, rel: str, body: bytes) -> bool:
        """
        OneDrive says the file changed since it was downloaded. Often only its
        version tag moved (the OneDrive app re-saving it) and the content is the
        same - then just save. If the content really changed and it is JSON,
        merge the two (keeping both sides' entries). Otherwise leave OneDrive's
        copy alone, note the conflict and carry on - never stop the whole run.
        """
        meta = requests.get(self._url(rel), headers=self._h(), timeout=60)
        cur = requests.get(self._url(rel, ":/content"), headers=self._h(), timeout=120)
        if not (meta.ok and cur.ok):
            self.conflicts.append(rel)
            return False
        etag = meta.json().get("eTag")
        if cur.content != self.hashes.get(rel, b"\0"):
            merged = merge_json(cur.content, body) if rel.endswith(".json") else None
            if merged is None:
                self.conflicts.append(rel)
                return False
            body = merged
        r = requests.put(self._url(rel, ":/content"),
                         headers=self._h({"Content-Type": "application/octet-stream", "If-Match": etag}),
                         data=body, timeout=120)
        if not r.ok:
            self.conflicts.append(rel)
            return False
        self.etags[rel] = r.json().get("eTag")
        self.hashes[rel] = body
        return True


def merge_json(theirs: bytes, mine: bytes):
    """
    Merge two versions of a store JSON file. Dictionaries are combined key by
    key (an entry only one side has is kept; for an entry both have, the one
    with the later last_active / updated date wins, else this run's). Lists of
    plain values are unioned. Returns bytes, or None if it can't be merged.
    """
    try:
        a, b = json.loads(theirs), json.loads(mine)
    except Exception:
        return None

    def stamp(x):
        return (x.get("last_active") or x.get("updated") or x.get("last_updated") or "") if isinstance(x, dict) else ""

    def m(x, y):
        if isinstance(x, dict) and isinstance(y, dict):
            out = dict(x)
            for k, v in y.items():
                if k not in out:
                    out[k] = v
                elif isinstance(out[k], dict) and isinstance(v, dict) and (stamp(out[k]) or stamp(v)):
                    out[k] = out[k] if stamp(out[k]) > stamp(v) else v
                else:
                    out[k] = m(out[k], v)
            return out
        if isinstance(x, list) and isinstance(y, list) and all(not isinstance(i, (dict, list)) for i in x + y):
            return x + [i for i in y if i not in x]
        return y

    return json.dumps(m(a, b), indent=2, ensure_ascii=False).encode("utf-8")
