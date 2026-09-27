"""
board_update.py - weekly board update email (Control Room: 'Weekly board update').

Now built on the knowledge store (see board_weekly.py): starts from the matters
that moved, follows up what you told the board last time, and learns from how
you edit each draft. The draft lands in Outlook Drafts, unaddressed, for you to
review, address and send.

  py board_update.py run [--days 7]      draft to Outlook + markdown copy
  py board_update.py preview [--days 7]  print it, create nothing
  py board_update.py setup / test        shared sign-in / connection check
"""
import sys
import argparse
import datetime
from pathlib import Path

from core.env import load_env
load_env()
from core.config import (BASE_DIR, STORE_DIR, COMPANY_NAME, MD_NAME, MD_TITLE, VOICE_RULES)
from core import graph
from core.llm import get_client

MAX_EMAILS_PER_FOLDER = 75
MAX_EMAILS_TOTAL      = 400
MAX_BODY_CHARS        = 1500


def cmd_run(args, preview=False):
    import board_weekly as bw
    days = args.days or 7
    print(f"\n  {COMPANY_NAME.upper()} - WEEKLY BOARD UPDATE  ({days}-day window)\n")
    token = graph.get_access_token()
    emails = graph.run_scan(token, days, max_per_folder=MAX_EMAILS_PER_FOLDER,
                            max_total=MAX_EMAILS_TOTAL, body_chars=MAX_BODY_CHARS,
                            label="BOARD UPDATE", refresh_folders=args.refresh_folders)
    instr = Path(STORE_DIR) / "instructions.md"
    instructions = instr.read_text(encoding="utf-8") if instr.exists() else ""
    client = get_client()
    store = Path(STORE_DIR)
    if preview:
        sent = bw.fetch_sent_updates(token, store)
        out = bw.generate(client, store, emails, days, sent, VOICE_RULES, instructions)
        print(bw.render_email(out["items"], MD_NAME, MD_TITLE, out.get("intro", ""))[1])
        return
    res = bw.make_update(token, client, store, emails, days, VOICE_RULES, COMPANY_NAME,
                         MD_NAME, MD_TITLE, instructions, say=lambda m: print("  " + m))
    md = BASE_DIR / f"board_update_{datetime.date.today():%Y-%m-%d}.md"
    md.write_text(res["text"], encoding="utf-8")
    print(f"\n  Draft is in Outlook Drafts: \"{res['subject']}\"")
    print(f"  Copy saved: {md.name}")
    print(f"  Review notes (sources, what was left out): state gas knowledge/board/drafts/{res['stamp']}.html\n")


def main():
    ap = argparse.ArgumentParser(description="Weekly board update")
    ap.add_argument("command", nargs="?", default="run",
                    choices=["run", "preview", "setup", "test"])
    ap.add_argument("--days", type=int, default=None)
    ap.add_argument("--refresh-folders", action="store_true")
    args = ap.parse_args()
    if args.command == "setup":
        graph.setup_auth()
    elif args.command == "test":
        tok = graph.get_access_token()
        print("  Connected." if tok else "  Not signed in - run setup.")
    else:
        cmd_run(args, preview=args.command == "preview")


if __name__ == "__main__":
    main()
