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
        self.etags[rel] = meta.json().get("eTag")
        data = requests.get(self._url(rel, ":/content"), headers=self._h(), timeout=120)
        data.raise_for_status()
        dest = self.local / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data.content)
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
            raise RuntimeError(f"{rel} changed in OneDrive during the run - not overwritten")
        r.raise_for_status()
        self.etags[rel] = r.json().get("eTag")
        self.hashes[rel] = body
        return True
