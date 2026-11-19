"""
levelbuilder_writer.py

Shared Levelbuilder read/write helpers for the Streamlit app. Same
Playwright mechanics already proven working in fill_existing_sections.py
and create_new_sections.py (live re-scanning before each group to stay
correct even as sections get created, preceding::/following:: field
lookups, the Callout modal flow) -- just without the pause-and-confirm
steps, since the app's own review screen is what stands in for those
now. Nothing here saves the lesson; that's always a manual click by
the person, in the browser window this leaves open.
"""

import json

from playwright.sync_api import sync_playwright


CALLOUT_TYPE_TO_VALUE = {
    "Teaching Tip": "teachingTip",
    "Assessment Opportunity": "assessmentOpportunity",
}


def get_live_blocks(page) -> list:
    """Reads the page's own hidden `activities` field fresh -- the real,
    current structure, confirmed to be a reliable read source (though a
    one-way mirror, not writable) earlier in this project."""
    raw = page.eval_on_selector('input[name="activities"]', "el => el.value")
    return json.loads(raw)


def find_absolute_index_by_section_id(live_blocks: list, target_id) -> int:
    """
    Finds the CURRENT 0-indexed position (among ALL "+Activity Section"
    buttons on the page) of the section with this id, from a fresh live
    scan -- correct even after earlier insertions in this run, or a
    section created but not yet saved (which has no real id yet;
    sec.get("id") skips those safely rather than crashing).
    """
    abs_idx = 0
    for block in live_blocks:
        for sec in block["activitySections"]:
            if sec.get("id") == target_id:
                return abs_idx
            abs_idx += 1
    raise ValueError(f"Section id {target_id} not found in the current live page state")


def _fill_common_fields(page, button_locator, sec: dict, direction: str, log) -> None:
    """
    Fills Title/Duration/Textarea/Remarks and adds any Callouts for one
    section, whose fields sit either BEFORE button_locator (an existing
    section's own trailing button -- direction="preceding") or AFTER it
    (a just-created section's fields -- direction="following").
    """
    axis = "preceding" if direction == "preceding" else "following"
    # NOTE: the Title input and the Progression Title input share the exact
    # same "width: 275px" styling, so searching for "nearest 275px input"
    # from the button is ambiguous on any ANCHOR section (one that already
    # has a Progression Title) -- confirmed via direct DOM inspection that
    # it can land on Progression Title instead of this section's own Title
    # field, silently overwriting it. The Duration field has no such
    # duplicate (still exactly one "width: 50px" input per section), and
    # the real Title input is always its immediate sibling -- same <label>,
    # laid out as `Title: [input] Duration: [input]` -- regardless of
    # section direction. So anchor on Duration first (unambiguous), then
    # step back to the 275px input immediately before THAT, rather than
    # before the button.
    duration_input = button_locator.locator(f"xpath={axis}::input[contains(@style,'width: 50px')][1]")
    title_input = duration_input.locator("xpath=preceding::input[contains(@style,'width: 275px')][1]")
    # Not filtering on rows='2' -- confirmed that attribute changes once
    # a textarea has real content in it (it auto-grows), which broke
    # this selector on every section after the first successful fill.
    # Just the nearest textarea in this direction is reliable instead.
    textarea = button_locator.locator(f"xpath={axis}::textarea[1]")
    remarks_checkbox = button_locator.locator(f"xpath={axis}::input[@type='checkbox'][1]")

    if sec["title"]:
        title_input.fill(sec["title"])
        actual = title_input.input_value()
        if actual != sec["title"]:
            log(
                f"  WARNING: Title field for slide {sec['slide']} shows "
                f"{actual!r} after filling, not the expected {sec['title']!r}. "
                f"The write may have silently failed -- check this field by hand."
            )
    if sec["duration"]:
        duration_input.fill(str(sec["duration"]))
        actual = duration_input.input_value()
        if actual != str(sec["duration"]):
            log(
                f"  WARNING: Duration field for slide {sec['slide']} shows "
                f"{actual!r} after filling, not the expected {sec['duration']!r}. "
                f"The write may have silently failed -- check this field by hand."
            )
    textarea.fill(sec["body_html"])
    actual = textarea.input_value()
    if actual.strip() != sec["body_html"].strip():
        log(
            f"  WARNING: textarea for slide {sec['slide']} doesn't match what "
            f"was written -- the write may have silently failed. Check this "
            f"section's content by hand before trusting it."
        )
    if sec["remarks_checked"]:
        remarks_checkbox.check()
        if not remarks_checkbox.is_checked():
            log(
                f"  WARNING: Remarks checkbox for slide {sec['slide']} isn't "
                f"actually checked after trying -- check this by hand."
            )

    for callout in sec["callouts"]:
        value = CALLOUT_TYPE_TO_VALUE.get(callout["type"])
        if value is None:
            log(f"  WARNING: callout type {callout['type']!r} has no known dropdown value -- skipped.")
            continue
        callout_button = button_locator.locator(f"xpath={axis}::button[contains(., 'Callout')][1]")
        callout_button.click()
        modal = page.locator(".modal.dash_modal")
        modal.locator("select").select_option(value=value)
        modal.locator("textarea[name='callout']").fill(callout["text"])
        modal.locator("button:has-text('Save and close')").click()


def run_fill_existing(page, all_operations: list, log=print) -> None:
    """
    Fills EXISTING (already-created) sections per the assignment plan.
    Re-checks the live page before each one, so this is safe to run
    before OR after run_create_new in the same session.
    """
    section_buttons = page.locator("button.add-activity-section")
    for op in all_operations:
        live_blocks = get_live_blocks(page)
        abs_idx = find_absolute_index_by_section_id(live_blocks, op["section_id"])
        button = section_buttons.nth(abs_idx)
        sec = op["slide"]
        _fill_common_fields(page, button, sec, "preceding", log)
        log(f"Filled existing section for slide {sec['slide']} ({op['block']})")


def run_create_new(page, all_groups: list, log=print) -> None:
    """
    Creates brand-new sections for leftover content per the assignment
    plan. Re-checks the live page once per GROUP (not once per slide --
    within a group, each new section's own trailing button is found via
    "following::" from the one just created, same as the already-proven
    chaining approach).
    """
    for group in all_groups:
        if group["insert_at_block_start"]:
            log(
                f"  SKIPPED group for {group['block']} -- insertion at the "
                f"very start of a block (no existing section to anchor on) "
                f"isn't supported yet. Create that section by hand first."
            )
            continue

        live_blocks = get_live_blocks(page)
        anchor_id = group["insert_after_section_id"] or group["insert_after_anchor_id"]
        abs_idx = find_absolute_index_by_section_id(live_blocks, anchor_id)

        section_buttons = page.locator("button.add-activity-section")
        current_button = section_buttons.nth(abs_idx)

        for sec in group["slides"]:
            current_button.click()
            _fill_common_fields(page, current_button, sec, "following", log)
            log(f"Created new section for slide {sec['slide']} ({group['block']})")
            current_button = current_button.locator(
                "xpath=following::button[contains(@class,'add-activity-section')][1]"
            )


def scan_levelbuilder_live(edit_url: str, session_file: str = "levelbuilder_session.json", browser_name: str = "firefox") -> list:
    """
    Opens a short-lived HEADLESS browser just to read the lesson's
    current structure (for the plan-review step) -- no pauses, no
    visible window, closes itself when done. Raises a clear error if
    the saved session can't actually reach the page (e.g. expired),
    rather than returning something that looks like an empty lesson.
    """
    with sync_playwright() as p:
        browser_type = getattr(p, browser_name)
        browser = browser_type.launch(headless=True)
        context = browser.new_context(storage_state=session_file)
        page = context.new_page()
        page.goto(edit_url)
        page.wait_for_load_state("networkidle")
        try:
            # "networkidle" doesn't guarantee React has actually finished
            # rendering the lesson -- wait for the real data field to
            # exist before reading it. A bigger lesson can genuinely take
            # a while to render, so this is generous on purpose; a real
            # expired-session case still fails fast (the field never
            # appears at all), it just won't be mistaken for a slow page.
            # state="attached" (NOT the default "visible") -- this is a
            # type="hidden" input, which can never pass a visibility
            # check no matter how long we wait. Confirmed via direct
            # browser inspection: the field is populated correctly from
            # the moment the page loads; the earlier default-state wait
            # was timing out on something that could never succeed.
            page.wait_for_selector('input[name="activities"]', state="attached", timeout=15000)
            raw_value = page.eval_on_selector('input[name="activities"]', "el => el.value")
        except Exception as e:
            current_url = page.url
            browser.close()
            if current_url.strip().rstrip("/") != edit_url.strip().rstrip("/"):
                raise RuntimeError(
                    f"Couldn't read the lesson's structure -- the page "
                    f"redirected to {current_url!r} instead of staying on "
                    f"the lesson, which usually means your saved session "
                    f"really has expired. Log in again and retry."
                ) from e
            raise RuntimeError(
                f"Couldn't read the lesson's structure after waiting 60 "
                f"seconds, even though the page stayed on the right URL "
                f"({current_url!r}) -- this looks more like a slow-loading "
                f"page than an expired session. Try again, or check your "
                f"internet connection."
            ) from e
        browser.close()
    return json.loads(raw_value)


def login_and_save_session(
    edit_url: str,
    session_file: str = "levelbuilder_session.json",
    browser_name: str = "firefox",
    timeout_seconds: int = 300,
) -> None:
    """
    Opens a real, visible browser for a one-time login, then WAITS --
    polling the page itself, within this one call -- until it detects
    you've actually reached the lesson edit page (by checking for the
    hidden `activities` field that only exists there, not on a login
    screen). Once detected, saves the session and closes.

    This deliberately does everything in a single call rather than
    splitting "open browser" and "save session" across two separate
    button clicks: Playwright's sync API isn't safe to use across
    different Streamlit reruns (each one runs in a different thread
    behind the scenes), so the browser/context/connection can't be
    kept alive in session_state between two clicks the way a simpler
    design might assume.
    """
    import time

    with sync_playwright() as p:
        browser_type = getattr(p, browser_name)
        browser = browser_type.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(edit_url)

        waited = 0
        poll_interval = 2
        logged_in = False
        while waited < timeout_seconds:
            try:
                logged_in = page.evaluate(
                    "() => !!document.querySelector('input[name=\"activities\"]')"
                )
            except Exception:
                logged_in = False
            if logged_in:
                break
            time.sleep(poll_interval)
            waited += poll_interval

        if not logged_in:
            browser.close()
            raise TimeoutError(
                "Didn't detect a successful login within the time limit. "
                "Try again, and make sure you reach the actual lesson "
                "edit page (not left on a login or redirect screen)."
            )

        context.storage_state(path=session_file)
        browser.close()
