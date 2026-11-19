#!/usr/bin/env python3
"""
scan_levelbuilder_lesson.py

READ-ONLY. Opens a Levelbuilder lesson edit page and reports the real,
current structure by reading the page's own hidden `activities` field --
a JSON snapshot of the whole lesson (confirmed present via browser
inspection) -- rather than guessing from visible DOM layout. This is
far more reliable than DOM-position heuristics (an earlier version of
this script broke on a stray trailing space in a field's value).

This never clicks, types, or saves anything -- it only reads a field's
current value. Safe to run any time.

NOTE: this hidden field has been confirmed to be a ONE-WAY SNAPSHOT --
editing it does not change what's on the page. So it's a reliable READ
source, but the write step still has to go through the real UI (typing
into fields, clicking buttons), same as what we already tested
separately in test_levelbuilder_write.py.
"""

import argparse
import json
from playwright.sync_api import sync_playwright


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--edit-url", required=True)
    parser.add_argument("--session-file", default="levelbuilder_session.json")
    parser.add_argument("--browser", default="firefox", choices=["firefox", "chromium", "webkit"])
    parser.add_argument("--out", default="lesson_scan.json", help="Where to save the raw JSON")
    args = parser.parse_args()

    with sync_playwright() as p:
        browser_type = getattr(p, args.browser)
        browser = browser_type.launch(headless=False)
        context = browser.new_context(storage_state=args.session_file)
        page = context.new_page()
        page.goto(args.edit_url)
        page.wait_for_load_state("networkidle")

        raw_value = page.eval_on_selector(
            'input[name="activities"]', "el => el.value"
        )
        browser.close()

    blocks = json.loads(raw_value)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(blocks, f, indent=2)
    print(f"Full raw data saved to {args.out}\n")

    for block in blocks:
        sections = block["activitySections"]
        print(f"=== {block['name']}  (duration: {block['duration']} mins, {len(sections)} sections) ===")
        for i, sec in enumerate(sections, start=1):
            num_levels = len(sec.get("scriptLevels", []))
            # A real anchor needs BOTH a Progression Title AND at least
            # one actual Level -- a Progression Title alone can be a
            # placeholder typed in before Levels were attached.
            is_anchor = bool(sec.get("progressionName")) and num_levels > 0
            kind = "ANCHOR" if is_anchor else "blank"
            prog = f" (Progression Title: {sec['progressionName']!r})" if sec.get("progressionName") else ""
            has_content = bool((sec.get("description") or "").strip())
            content_state = "HAS CONTENT" if has_content else "empty"
            num_tips = len(sec.get("tips", []))
            print(
                f"  [{i}] {kind}{prog} -- {num_levels} level(s) -- "
                f"textarea: {content_state} -- callouts: {num_tips}"
            )
            if has_content:
                preview = sec["description"][:80]
                print(f"        preview: {preview!r}...")
        print()


if __name__ == "__main__":
    main()
