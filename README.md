# State Gas Tools — Consolidated Suite

Refactor of `inbox_actions.py`, `board_update.py`, `board_briefing.py` (+ context
module), `board_report_analyser.py`, and `shared_knowledge.py` into one package
with shared infrastructure and a unified memory layer.

## Layout

```
stategas_tools/
  core/
    config.py        paths, company identity, strategic pillars, shared voice rules
    env.py           .env discovery
    graph.py         auth (one scope set), folder discovery (cached), email fetch
    llm.py           model config, retry wrapper, robust JSON parser
    knowledge.py     CompanyKnowledge / InboxState / PillarHistory, briefing loader
  inbox_actions.py   weekly action extraction → review dashboard → MS To Do
  board_update.py    weekly board update email (markdown + Outlook draft)
  board_briefing.py  monthly briefing → board_report_data.json (PPTX) + markdown, --docx
  docx_export.py     Word-draft renderer (ported from the retired analyser)
  migrate_knowledge.py   one-off data migration (run first — see below)
```

`board_report_analyser.py`, `shared_knowledge.py`, and
`board_briefing_context.py` are **retired** — do not copy them into this folder.

## Getting started

1. Copy your existing `.env`, `.outlook_token_cache.bin`, `inbox_state.json`,
   `context_briefing.txt`, and the old knowledge JSONs (`shared_knowledge.json`,
   `board_context.json`, `board_briefing_context.json`) into this folder.
2. Run the migration (once):

   ```
   py migrate_knowledge.py --dry-run    # preview
   py migrate_knowledge.py              # write company_knowledge.json + pillar_history.json
   ```

3. Re-auth once with the unified scope set (the old cache may lack Tasks.ReadWrite
   consent depending on which tool created it):

   ```
   py inbox_actions.py setup
   ```

4. Normal use:

   ```
   py inbox_actions.py run [--days 14]
   py board_update.py run [--days 7]
   py board_briefing.py run [--days 30] [--docx]
   ```

   `--days` skips the interactive prompt (needed for scheduled/remote runs).
   `--refresh-folders` forces folder rediscovery (otherwise cached 7 days).
   `BOARD_TOOLS_MODEL` in `.env` overrides the Claude model for every call.

## What changed (fixes from the code review)

| Ref | Fix |
|-----|-----|
| B1 | Pushing tasks to To Do records `pushed_to_todo` in inbox state again — push moved back server-side |
| S1 | The dashboard HTML no longer contains your Graph access token |
| B2 | Briefing markdown/context update now read the fields the synthesis actually produces |
| B3 | Pillar RAG history persists (`pillar_history.json`) — next month's run sees last month's status |
| B4 | Knowledge prompt blocks inject newest-first with per-section budgets instead of truncating the newest away |
| B5 | `company_knowledge.json` is pruned (active→dormant after 60d; closed removed after 90d; dormant after 365d) |
| B6 | Report headers show the lookback you actually chose |
| B7 | Email fetch follows pagination — busy folders no longer drop messages silently |
| B8 | `context_briefing.txt` is read head+tail, so auto-appended items stay visible |
| B9 | One scope set for the shared token cache |
| B10 | Dead code removed (`serve_dashboard` old path, `update-context`, dormant section, unused imports) |
| B11 | Decision-log matching handles RE:/FW: subject prefixes |
| P1 | Bodies fetched in the folder list call (one request per page, not per email) — ~10x faster scans |
| P2 | Folder discovery cached for 7 days |
| P3 | Fixed 30s/20s sleeps removed; retry honours the API's retry-after header |
| P4 | Company knowledge loaded once per run, not per batch |
| — | Post-run maintenance is ONE Claude call per tool (was up to three) |
| — | Single `merge_matters`, single JSON parser, single retry wrapper, shared voice rules |

## Behaviour changes to be aware of

- **Dashboard push needs the local server running.** Opening the saved
  `inbox_dashboard.html` directly from disk shows the review but disables
  Push/Resolve — run `py inbox_actions.py review` to serve it. This is the
  price of not writing your mail/tasks token into an HTML file.
- **`board_context.json` is no longer read or written.** Everything cross-cutting
  lives in `company_knowledge.json`. The board update's post-run maintenance is
  one patch call instead of two.
- **`board_briefing.py --days` default is 30**; `build-context` defaults to 90.
- The PPTX data file (`board_report_data.json`) format is **unchanged** —
  `generate_board_deck.js` works as before.

## Scheduling / remote runs (GitHub Actions ready)

Every pipeline is now non-interactive when `--days` is supplied. For remote
runs: schedule `inbox_actions.py analyse --days 7`, `board_update.py run --days 7`,
and `board_briefing.py run --days 30` via cron/Actions, persist
`.outlook_token_cache.bin` between runs, and keep the interactive review step
(`inbox_actions.py review`) local.
