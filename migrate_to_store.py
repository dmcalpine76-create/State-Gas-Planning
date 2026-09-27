"""
migrate_to_store.py - one-off move of the planning tools' memory into the
shared knowledge store ("state gas knowledge" beside this folder).

  * company_knowledge.json  -> matters.json + people.json + facts.json
  * inbox_state.json (this folder, 33 runs to 31 Aug) merged with the copy in
    the morning briefing folder (5 runs, 7-21 Sep) -> inbox/inbox_state.json
  * pillar_history.json     -> pillars.json
  * originals copied untouched to archive/pre-migration-<date>/

Contents are moved as-is: nothing is trimmed or rewritten, so inbox_actions
sees exactly the same knowledge it saw before. Safe to re-run with --force.

  py migrate_to_store.py            (refuses if the store already has data)
  py migrate_to_store.py --dry-run
"""
import sys, json, shutil, datetime, re
from pathlib import Path

HERE  = Path(__file__).resolve().parent
STORE = HERE.parent / "state gas knowledge"
MB    = HERE.parent / "morning briefing system"
DRY   = "--dry-run" in sys.argv
FORCE = "--force" in sys.argv

def load(p):
    return json.loads(Path(p).read_text(encoding="utf-8")) if Path(p).exists() else {}

def dump(p, data):
    if DRY:
        print(f"   (dry) would write {p.relative_to(STORE)}"); return
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)
    print(f"   wrote {p.relative_to(STORE)}")

def newer(a, b, key):
    return a if (a.get(key) or "") >= (b.get(key) or "") else b

def merge_inbox(a, b):
    """a = planning tools copy (primary), b = morning briefing copy."""
    out = {
        "last_updated": max(a.get("last_updated", ""), b.get("last_updated", "")),
        "run_count":    a.get("run_count", 0) + b.get("run_count", 0),
        "topics": dict(a.get("topics", {})),
        "key_people": dict(a.get("key_people", {})),
    }
    for k, v in b.get("topics", {}).items():
        out["topics"][k] = newer(out["topics"][k], v, "last_seen") if k in out["topics"] else v
    for k, v in b.get("key_people", {}).items():
        out["key_people"][k] = newer(out["key_people"][k], v, "last_seen") if k in out["key_people"] else v

    # Both copies numbered their actions from act-001: renumber the second
    # copy after the first's highest id so resolve-by-id can't hit the wrong one.
    def num(x):
        m = re.match(r"act-(\d+)$", x.get("id", "") or "")
        return int(m.group(1)) if m else 0
    top = max([num(x) for x in a.get("open_actions", []) + a.get("closed_actions", [])] or [0])
    remap = {}
    for x in b.get("open_actions", []) + b.get("closed_actions", []):
        top += 1
        remap[x.get("id")] = f"act-{top:03d}"
    def rn(lst):
        res = []
        for x in lst:
            y = dict(x); y["id"] = remap.get(x.get("id"), x.get("id")); res.append(y)
        return res
    seen = set()
    def dedupe(lst):
        res = []
        for x in lst:
            t = (x.get("title") or "").strip().lower()
            if t and t in seen: continue
            seen.add(t); res.append(x)
        return res
    out["open_actions"]   = dedupe(a.get("open_actions", []) + rn(b.get("open_actions", [])))
    out["closed_actions"] = dedupe(a.get("closed_actions", []) + rn(b.get("closed_actions", [])))
    return out

def main():
    if not STORE.exists():
        sys.exit(f"Store folder not found: {STORE}")
    if (STORE / "matters.json").exists() and not (FORCE or DRY):
        sys.exit("Store already populated - re-run with --force to overwrite.")

    ck  = load(HERE / "company_knowledge.json")
    ia  = load(HERE / "inbox_state.json")
    ib  = load(MB / "inbox_state.json")
    ph  = load(HERE / "pillar_history.json")

    stamp = datetime.date.today().isoformat()
    bk = STORE / "archive" / f"pre-migration-{stamp}"
    if not DRY:
        bk.mkdir(parents=True, exist_ok=True)
        for src, name in [(HERE / "company_knowledge.json", "company_knowledge.json"),
                          (HERE / "inbox_state.json", "inbox_state.planning-tools.json"),
                          (MB / "inbox_state.json", "inbox_state.morning-briefing.json"),
                          (HERE / "pillar_history.json", "pillar_history.json"),
                          (HERE / "context_briefing.txt", "context_briefing.txt")]:
            if src.exists(): shutil.copy2(src, bk / name)
        print(f"   originals copied to archive/{bk.name}/")

    dump(STORE / "matters.json", {"matters": ck.get("matters", {})})
    dump(STORE / "people.json",  {"key_people": ck.get("key_people", {})})
    dump(STORE / "facts.json",   {"last_updated": ck.get("last_updated", ""),
                                  "company_profile": ck.get("company_profile", ""),
                                  "facts": ck.get("facts", [])})
    merged = merge_inbox(ia, ib)
    dump(STORE / "inbox" / "inbox_state.json", merged)
    dump(STORE / "pillars.json", ph)
    if not (STORE / "deadlines.json").exists():
        dump(STORE / "deadlines.json", {"deadlines": []})

    print(f"\n   matters {len(ck.get('matters', {}))}, people {len(ck.get('key_people', {}))}, "
          f"facts {len(ck.get('facts', []))}")
    print(f"   inbox: topics {len(merged['topics'])}, people {len(merged['key_people'])}, "
          f"open {len(merged['open_actions'])}, closed {len(merged['closed_actions'])}, "
          f"runs {merged['run_count']}")

if __name__ == "__main__":
    main()
