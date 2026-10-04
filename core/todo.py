"""
core/todo.py - one reader for ALL of the MD's Microsoft To Do lists, shared by
the Week planner, inbox actions and the board tools, plus the two-way link
between inbox actions' Resolve and To Do completion.

  fetch_all_tasks(token)        every task in every list (open, and completed
                                in the last N days), each tagged with its list
  complete_matching(token, ts)  tick off open To Do tasks whose title matches
  sync_completed(state, token)  resolve inbox actions whose To Do task is done
"""
import re
import datetime

GRAPH = "https://graph.microsoft.com/v1.0"


def _get(token, url, params=None):
    import requests
    r = requests.get(url if url.startswith("http") else GRAPH + url,
                     headers={"Authorization": f"Bearer {token}"}, params=params, timeout=60)
    r.raise_for_status()
    return r.json()


def _norm(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (title or "").lower()).strip()


def all_lists(token) -> list:
    out, url = [], "/me/todo/lists"
    while url:
        d = _get(token, url)
        out += d.get("value", [])
        url = d.get("@odata.nextLink")
    return out


def fetch_all_tasks(token, completed_days: int = 14, per_list_cap: int = 500) -> list:
    """Open tasks in every list, plus tasks completed in the last completed_days."""
    since = (datetime.date.today() - datetime.timedelta(days=completed_days)).isoformat()
    tasks = []
    for lst in all_lists(token):
        url, params, n = f"/me/todo/lists/{lst['id']}/tasks", {"$top": "100"}, 0
        while url and n < per_list_cap:
            d = _get(token, url, params)
            params = None
            for t in d.get("value", []):
                n += 1
                done = ((t.get("completedDateTime") or {}).get("dateTime") or "")[:10]
                if t.get("status") == "completed" and done < since:
                    continue
                tasks.append({
                    "id":         t.get("id"),
                    "list_id":    lst["id"],
                    "list":       lst.get("displayName", ""),
                    "title":      (t.get("title") or "").strip(),
                    "status":     t.get("status"),
                    "importance": t.get("importance"),
                    "created":    (t.get("createdDateTime") or "")[:10],
                    "modified":   (t.get("lastModifiedDateTime") or "")[:10],
                    "completed":  done,
                    "due":        ((t.get("dueDateTime") or {}).get("dateTime") or "")[:10],
                    "note":       ((t.get("body") or {}).get("content") or "")[:300],
                })
            url = d.get("@odata.nextLink")
    return tasks


def complete_matching(token, titles: list) -> int:
    """Mark open To Do tasks complete where the title matches one given. Returns count."""
    import requests
    want = {_norm(t) for t in titles if t}
    if not want:
        return 0
    done = 0
    for t in fetch_all_tasks(token, completed_days=0):
        if t["status"] != "completed" and _norm(t["title"]) in want:
            r = requests.patch(f"{GRAPH}/me/todo/lists/{t['list_id']}/tasks/{t['id']}",
                               headers={"Authorization": f"Bearer {token}",
                                        "Content-Type": "application/json"},
                               json={"status": "completed"}, timeout=30)
            if r.ok:
                done += 1
    return done


def sync_completed(state, token) -> int:
    """Resolve inbox actions whose matching To Do task has been completed."""
    tasks = fetch_all_tasks(token, completed_days=60)
    completed = {_norm(t["title"]) for t in tasks if t["status"] == "completed"}
    still_open = {_norm(t["title"]) for t in tasks if t["status"] != "completed"}
    ids = [a.get("id") for a in state.data.get("open_actions", [])
           if _norm(a.get("title", "")) in completed and _norm(a.get("title", "")) not in still_open]
    if ids:
        state.mark_resolved(ids)
    return len(ids)
