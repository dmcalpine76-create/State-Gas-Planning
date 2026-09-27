"""
The tools the Control Room may start on this PC. Re-read on every request,
so adding one here needs no launcher restart. {py} = python, {days} = the
lookback the page sent (checked to be a whole number from 1 to 120).
"""
TOOLS = {
    "inbox":          {"title": "Inbox actions",
                       "cmd": '"{py}" inbox_actions.py run --days {days}'},
    "inbox_review":   {"title": "Inbox actions review",
                       "cmd": '"{py}" inbox_actions.py review'},
    "board_update":   {"title": "Weekly board update",
                       "cmd": '"{py}" board_update.py run --days {days}'},
    "board_briefing": {"title": "Monthly board pack",
                       "cmd": '"{py}" board_briefing.py run --days {days} --docx'
                              ' && node generate_board_deck.js --data board_report_data.json'},
    "diary":          {"title": "Plan diary blocks",
                       "cmd": '"{py}" outlook_scheduler.py dashboard'},
    "weekly_setup":   {"title": "Sign in for the Sunday run",
                       "cmd": '"{py}" weekly_plan.py setup'},
    "weekly_local":   {"title": "Weekly plan (on this PC)",
                       "cmd": '"{py}" weekly_plan.py run --days {days}'},
}
