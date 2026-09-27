"""
core/config.py — single source of truth for paths, company identity,
and the State Gas strategic pillar definitions.

Everything that used to be copy-pasted across four scripts lives here.
"""

from pathlib import Path

# Package root — the folder containing inbox_actions.py / board_update.py etc.
BASE_DIR = Path(__file__).resolve().parent.parent

# ── Company identity ─────────────────────────────────────────────────────────
COMPANY_NAME   = "State Gas Limited"
COMPANY_TICKER = "ASX: GAS"
MD_NAME        = "Doug McAlpine"
MD_TITLE       = "Managing Director"

# ── Memory / state files ─────────────────────────────────────────────────────
# company_knowledge.json replaces BOTH shared_knowledge.json and
# board_context.json (run migrate_knowledge.py once to merge them).
COMPANY_KNOWLEDGE_FILE = BASE_DIR / "company_knowledge.json"
INBOX_STATE_FILE       = BASE_DIR / "inbox_state.json"
PILLAR_HISTORY_FILE    = BASE_DIR / "pillar_history.json"

# ── Shared knowledge store ───────────────────────────────────────────────────
# One folder in OneDrive shared by every State Gas planning process:
#   matters.json, people.json, facts.json   company knowledge (split by kind)
#   inbox/inbox_state.json                  inbox_actions' own memory
#   pillars.json                            monthly pillar RAG history
# Found beside this folder ("state gas knowledge"), or wherever
# STATEGAS_STORE_DIR points (cloud runs). If it is missing the tools fall
# back to the old files in this folder and say so loudly - see STORE_ACTIVE.
import os as _os
STORE_DIR = Path(_os.environ.get("STATEGAS_STORE_DIR")
                 or (BASE_DIR.parent / "state gas knowledge"))
STORE_ACTIVE = (STORE_DIR / "matters.json").exists()
if STORE_ACTIVE:
    INBOX_STATE_FILE    = STORE_DIR / "inbox" / "inbox_state.json"
    PILLAR_HISTORY_FILE = STORE_DIR / "pillars.json"
else:
    print(f"  !!  Knowledge store not found at {STORE_DIR}")
    print(f"      Using the old files in {BASE_DIR.name} - runs will NOT be shared.")
CONTEXT_BRIEFING_FILE  = BASE_DIR / "context_briefing.txt"

# ── Scheduling ───────────────────────────────────────────────────────────────
# The single Microsoft To Do list that is the scheduler's work queue.
# inbox_actions.py pushes approved actions INTO this list; outlook_scheduler.py
# reads tasks OUT of it. Defined here rather than in scheduling_rules.json
# because the rules dashboard rewrites global.* wholesale on save and would
# silently drop an extra key.
TASK_LIST_NAME = "Daily Priorities"

SCHEDULING_RULES_FILE = BASE_DIR / "scheduling_rules.json"
SCHEDULING_PLAN_FILE  = BASE_DIR / "scheduling_plan.json"
DURATION_HISTORY_FILE = BASE_DIR / ".duration_history.json"

# Brisbane — no DST, so a fixed offset is safe.
AEST_OFFSET_HOURS = 10

# ── Auth / caches ────────────────────────────────────────────────────────────
TOKEN_CACHE_FILE  = BASE_DIR / ".outlook_token_cache.bin"
FOLDER_CACHE_FILE = BASE_DIR / ".folder_cache.json"
FOLDER_CACHE_TTL_DAYS = 7

# ── Company specifics: voice, strategic pillars ──────────────────────────────
# Kept in the knowledge store (company_config.json), not here, so this code
# can live in a public repository. Generic fallbacks apply only if the store
# is missing.
import json as _json
_GENERIC_VOICE_RULES = """WRITING VOICE - apply to every matter you write:
1. FIRST PERSON. Write as the Managing Director - "I met with...", "My view is...".
2. INCLUDE JUDGMENT AND OPINION, not just facts.
3. FORWARD FOCUS. Every matter says what happens next and when.
4. SHORT AND DIRECT. 2-4 sentences per matter unless genuinely significant.
5. NO PROCEDURAL DETAIL unless the process IS the news.
6. CONVERSATIONAL BUT PROFESSIONAL."""
try:
    _CO = _json.loads((STORE_DIR / "company_config.json").read_text(encoding="utf-8"))
except (OSError, ValueError):
    _CO = {}
    print(f"  !!  company_config.json not found in {STORE_DIR} - using generic defaults")
VOICE_RULES       = _CO.get("voice_rules", _GENERIC_VOICE_RULES)
STRATEGIC_PILLARS = _CO.get("strategic_pillars", [])
PILLAR_ID_MAP     = _CO.get("pillar_id_map", {})
