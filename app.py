"""
app.py

Run with: streamlit run app.py

Walks through: paste the Slides URL and the Levelbuilder edit URL,
review the generated plan (resolving any titled-sub-activity-with-no-
Levels mismatches right here instead of needing a command-line flag),
approve it, and the actual write happens unattended -- no per-step
confirmation like the terminal scripts required while we were still
finding bugs. A visible browser window is left open at the end for
you to review and save yourself; nothing here ever clicks Save.
"""

import os
import re
import traceback

import streamlit as st

from speaker_notes_sync import (
    extract_activity_plan,
    download_pptx_via_drive_api,
    extract_drive_file_id,
    SECTION_ORDER,
)
from build_assignment_plan import (
    compute_fill_operations,
    compute_create_operations,
    build_plan_for_block,
    format_content_preview,
    is_real_anchor,
)
from project_assignment_plan import (
    extract_project_plan,
    build_project_plan,
    iter_day_sub_blocks,
    SUB_SECTIONS,
)
from levelbuilder_writer import (
    scan_levelbuilder_live,
    login_and_save_session,
    run_fill_existing,
    run_create_new,
)
from slide_images import fetch_slide_thumbnails

st.set_page_config(page_title="Levelbuilder Speaker Notes Tool", layout="wide")
st.title("Speaker Notes → Levelbuilder")

CREDENTIALS_PATH = "credentials.json"
TOKEN_PATH = "token.json"
SESSION_FILE = "levelbuilder_session.json"


def render_review_summary(all_fill_ops: list, all_create_groups: list, plan_markdown: list, mismatch_count: int) -> None:
    """
    A short, readable summary shown above the raw plan -- team feedback
    was that surfacing the whole plan_markdown dump by default was
    overwhelming, so now it's three counts anyone can scan in a second,
    plus the raw plan tucked into a collapsed "show details" expander
    instead of auto-expanded.
    """
    num_new_sections = sum(len(g["slides"]) for g in all_create_groups)
    col1, col2, col3 = st.columns(3)
    col1.metric("Sections to fill", len(all_fill_ops))
    col2.metric("New sections to create", num_new_sections)
    col3.metric("Needs your input", mismatch_count)

    if mismatch_count == 0:
        st.success("Everything lines up -- every titled sub-activity in the deck matches a Level section already in Levelbuilder.")
    else:
        st.info("Resolve the item(s) above before approving -- they decide where some content ends up.")

    with st.expander("Show full plan details"):
        st.markdown("\n".join(plan_markdown))


def render_anchor_resolution(block_label: str, candidate_titles: list) -> tuple:
    """
    Shown when a block's titled-sub-activity count doesn't match how
    many sections already have Levels attached in Levelbuilder -- i.e.
    the tool can't tell by itself which title(s) are missing a Level
    vs. which are just plain narrative text with its own heading.

    Replaced an earlier checkbox phrased as a negative ("doesn't need
    its own Level") -- team feedback was that it was still confusing
    to parse (a checked box meaning "this does NOT have something"
    takes an extra mental flip). This asks the same question as a
    plain either/or choice instead.

    IMPORTANT, clarified after a real question about the wording: this
    is asking about the named part OF THE DECK's speaker notes (the
    content under that heading) -- NOT "does a Levelbuilder section
    titled this already exist somewhere." Deck titles are matched to
    Levelbuilder sections purely by POSITION/ORDER, never by name. The
    earlier phrasing spelled that mechanic out in the UI itself (plus
    raw counts), which turned out to just be more wordiness to wade
    through for the same decision -- cut down to one short line per
    item now, phrased as a plain, self-contained question that doesn't
    need the preamble to make sense.

    The radio's default used to be "Yes" -- which meant every question
    counted as "answered" the instant it rendered, before anyone had
    actually looked at it, so the "Needs your input" count in
    render_review_summary could never drop to 0 no matter what you
    clicked (confirmed as a real bug, not just stale wording). Fixed by
    passing index=None, which leaves BOTH options unselected until the
    user actually clicks one -- no third "placeholder" option cluttering
    the row (tried that first; team feedback was it looked odd sitting
    there forever on the far left). Requires Streamlit 1.29+ (released
    Nov 2023), which "pip install streamlit" already gets by default.

    Returns (no_anchor, still_unanswered_count): no_anchor is the
    subset of candidate_titles marked as NOT needing a Level of their
    own (fed into no_anchor_titles, same as before); still_unanswered_count
    is how many of these questions haven't been clicked yet.
    """
    st.warning(f"**{block_label}**: a few activity sections here need a quick check:")
    no_anchor = set()
    still_unanswered = 0
    for title in candidate_titles:
        choice = st.radio(
            f'Does "{title}" need its own Level in Levelbuilder?',
            ["Yes", "No -- just text, no Level"],
            index=None,
            key=f"anchor_choice::{block_label}::{title}",
            horizontal=True,
        )
        if choice is None:
            still_unanswered += 1
        elif choice.startswith("No"):
            no_anchor.add(title)
    return no_anchor, still_unanswered


def get_or_fetch_thumbnails(slide_numbers: list) -> dict:
    """
    Returns {slide_number: thumbnail_image_bytes} for the given slide
    numbers, fetching only the ones not already cached this session
    (st.session_state["slide_thumbnails"]). These are actual downloaded
    image bytes (see slide_images.fetch_slide_thumbnails), not a live
    URL -- so unlike an early version of this, there's nothing here
    that can expire or get rate-limited later just by being displayed;
    once it's fetched once, it's good for the rest of the session.
    Returns {} (and shows nothing) when no Slides URL was used this run
    (a manually-exported .pptx has no presentation_id to look thumbnails
    up with).
    """
    presentation_id = st.session_state.get("presentation_id")
    if not presentation_id:
        return {}
    cache = st.session_state.setdefault("slide_thumbnails", {})
    missing = [n for n in slide_numbers if n not in cache]
    if missing:
        try:
            cache.update(fetch_slide_thumbnails(presentation_id, CREDENTIALS_PATH, TOKEN_PATH, missing))
        except Exception as e:
            st.warning(f"Couldn't load slide image(s): {e}")
    return {n: cache[n] for n in slide_numbers if n in cache}


def plan_slide_numbers(all_fill_ops: list, all_create_groups: list) -> list:
    """Every distinct slide number the plan actually references, across
    both fill and create operations -- the set of slides worth having
    an image for."""
    return sorted({op["slide"]["slide"] for op in all_fill_ops} | {
        sec["slide"] for group in all_create_groups for sec in group["slides"]
    })


def preload_plan_thumbnails(all_fill_ops: list, all_create_groups: list) -> None:
    """
    Fetches (and caches) the slide image for every slide the plan
    references, right as the plan itself is assembled -- NOT gated
    behind a separate "Load" button anymore. Confirmed this matters:
    Streamlit can't respond to ANY click, including Approve, while a
    script rerun is still in progress, so a LATER "Load slide images"
    click was effectively freezing the whole page (Approve included)
    for however long that fetch took, not just the image section it
    looked like it belonged to. Doing it here means there's only ever
    ONE wait, right after Analyze, instead of a second unrelated one
    later that happens to block an unrelated button.
    """
    if not st.session_state.get("presentation_id"):
        return
    needed = plan_slide_numbers(all_fill_ops, all_create_groups)
    if needed:
        with st.spinner(f"Loading {len(needed)} slide image(s) for review..."):
            get_or_fetch_thumbnails(needed)


def render_plan_items_with_images(all_fill_ops: list, all_create_groups: list) -> None:
    """
    Walks the SAME structured fill/create operations used to actually
    write Levelbuilder (not a re-parse of the markdown plan), showing
    each one's slide image right next to the content it's about to
    place -- so you can check the tool's work slide by slide without a
    third window open for the deck itself. Images are already cached by
    preload_plan_thumbnails by the time this renders, so this is just
    display -- no button, no wait here.
    """
    if not st.session_state.get("presentation_id"):
        st.caption("Slide images aren't available this run -- a manually-exported .pptx was used instead of a Slides URL.")
        return

    thumbnails = get_or_fetch_thumbnails(plan_slide_numbers(all_fill_ops, all_create_groups))

    for op in all_fill_ops:
        sec = op["slide"]
        col1, col2 = st.columns([1, 4])
        with col1:
            image_bytes = thumbnails.get(sec["slide"])
            if image_bytes:
                st.image(image_bytes, width=140)
        with col2:
            st.markdown(f"**Fill** existing section in *{op['block']}* ← slide {sec['slide']}")
            st.code(format_content_preview(sec), language=None)

    for group in all_create_groups:
        for sec in group["slides"]:
            col1, col2 = st.columns([1, 4])
            with col1:
                image_bytes = thumbnails.get(sec["slide"])
                if image_bytes:
                    st.image(image_bytes, width=140)
            with col2:
                st.markdown(f"**Create new section** in *{group['block']}* ← slide {sec['slide']}")
                st.code(format_content_preview(sec), language=None)


_ICON_TAG_RE = re.compile(r"<i\b[^>]*>\s*</i>\s*")


def format_preview_for_display(sec: dict) -> str:
    """
    A read-friendly version of a section's extracted content for the
    Browse-all-slides view -- strips the literal <i class="..."></i>
    icon tags (meaningless visual clutter outside of what Levelbuilder
    itself renders them as) and renders via st.markdown so long lines
    wrap instead of requiring horizontal scrolling, unlike the no-wrap
    monospace code block used in the plan-item view (which intentionally
    shows the byte-for-byte content that will actually get written).
    """
    text = _ICON_TAG_RE.sub("", sec["body_html"]).strip()
    if sec["callouts"]:
        text += "\n\n*Callout(s): " + ", ".join(c["type"] for c in sec["callouts"]) + "*"
    return text


def flatten_all_sections(plan: dict, active_mode: str) -> list:
    """
    All extracted slide sections across the whole deck, in slide order
    -- regardless of which Warm Up / Activity / Wrap Up (or Day) bucket
    each one ended up in, since a bucket's slides aren't necessarily
    contiguous with the next bucket's. Used only by the slide browser
    below, which is a plain "flip through the deck" view independent of
    the plan/mismatch-resolution above it.
    """
    sections = []
    if active_mode == "Multi-Day Project":
        for subsections in plan.values():
            for sub_name in SUB_SECTIONS:
                sections.extend(subsections[sub_name])
    else:
        for block_name in SECTION_ORDER:
            sections.extend(plan[block_name])
    return sorted(sections, key=lambda s: s["slide"])


def _render_slide_browser_nav(key_suffix: str, page: int, num_pages: int, total: int, page_size: int) -> int:
    """
    One Previous/Next row -- called once above the slide list and once
    below it, so you never have to scroll back up just to turn the
    page. Each call needs its own widget keys (key_suffix) since
    Streamlit requires unique keys even for two buttons that do the
    same thing. Returns the page number AFTER applying whichever button
    (if any) was clicked this rerun.
    """
    nav1, nav2, nav3 = st.columns([1, 2, 1])
    with nav1:
        if st.button("← Previous", disabled=(page == 0), key=f"slide_browser_prev_{key_suffix}"):
            page -= 1
    with nav3:
        if st.button("Next →", disabled=(page >= num_pages - 1), key=f"slide_browser_next_{key_suffix}"):
            page += 1
    with nav2:
        start, end = page * page_size + 1, min((page + 1) * page_size, total)
        st.caption(f"Slides {start}-{end} of {total} (page {page + 1} of {num_pages})")
    return page


def render_slide_browser(plan: dict, active_mode: str) -> None:
    """
    A standalone, paginated "browse the whole deck" view -- every slide
    in order with its extracted speaker notes and (on request) its
    thumbnail. Independent of the plan above: this is for general
    reference while reviewing, not tied to what the tool is actually
    about to write. Paginated at 8 slides/page so flipping through a
    full deck never fetches more than 8 thumbnails for a single page.

    Fetches each page's images automatically as you navigate -- no
    "Load" button (that was freezing the whole page, Approve button
    included, for however long the fetch took; see
    preload_plan_thumbnails). Bounded to 8 slides/page keeps any one
    page's wait small regardless.
    """
    all_sections = flatten_all_sections(plan, active_mode)
    if not all_sections:
        return

    PAGE_SIZE = 8
    num_pages = (len(all_sections) - 1) // PAGE_SIZE + 1
    page = max(0, min(st.session_state.get("slide_browser_page", 0), num_pages - 1))

    page = _render_slide_browser_nav("top", page, num_pages, len(all_sections), PAGE_SIZE)
    st.session_state["slide_browser_page"] = page

    page_sections = all_sections[page * PAGE_SIZE: (page + 1) * PAGE_SIZE]
    thumbnails = {}
    if not st.session_state.get("presentation_id"):
        st.caption("Slide images aren't available this run -- a manually-exported .pptx was used instead of a Slides URL.")
    else:
        with st.spinner("Loading slide images..."):
            thumbnails = get_or_fetch_thumbnails([s["slide"] for s in page_sections])

    for sec in page_sections:
        col1, col2 = st.columns([1, 4])
        with col1:
            image_bytes = thumbnails.get(sec["slide"])
            if image_bytes:
                st.image(image_bytes, width=160)
            else:
                st.caption(f"Slide {sec['slide']}")
        with col2:
            title_line = f"**{sec['title']}** ({sec['duration']} mins)" if sec["title"] else f"Slide {sec['slide']}"
            st.markdown(title_line)
            st.markdown(format_preview_for_display(sec))
        st.divider()

    page = _render_slide_browser_nav("bottom", page, num_pages, len(all_sections), PAGE_SIZE)
    st.session_state["slide_browser_page"] = page


# ----------------------------------------------------------------------
# Step 1: inputs
# ----------------------------------------------------------------------

mode = st.radio(
    "What are you syncing?",
    ["Single Lesson", "Multi-Day Project"],
    key="mode_choice",
    horizontal=True,
    help=(
        "Multi-Day Project is for Connect Task-style lessons where "
        "Levelbuilder groups sections under 'Day 1' / 'Day 2' / ... "
        "blocks (each containing its own Warm Up / Activity / Wrap Up "
        "divider sections), instead of the usual single Warm Up / "
        "Activity / Wrap Up blocks."
    ),
)
expected_days = None
if mode == "Multi-Day Project":
    expected_days = st.number_input(
        "How many days does this project span?", min_value=2, max_value=10, value=3, step=1,
        help="Used as a sanity check against how many 'Warm Up' markers the deck actually has -- not the thing that decides the split itself.",
    )

col1, col2 = st.columns(2)
with col1:
    slides_url = st.text_input("Google Slides URL", key="slides_url")
with col2:
    edit_url = st.text_input("Levelbuilder edit page URL", key="edit_url")

with st.expander("Deck too large to download directly? Use a manually-exported file instead"):
    st.caption(
        "Google's export API has a hard size limit that image-heavy decks "
        "can exceed, even when the file itself isn't unreasonably large. "
        "If Analyze fails with an 'exportSizeLimitExceeded' error, export "
        "the deck yourself (File → Download → Microsoft PowerPoint) and "
        "paste the saved file's path below."
    )
    pptx_override = st.text_input("Path to manually-exported .pptx file", key="pptx_override")

analyze_clicked = st.button(
    "Analyze", type="primary", disabled=not ((slides_url or pptx_override) and edit_url)
)

if analyze_clicked:
    st.session_state.pop("plan", None)
    st.session_state.pop("scanned_by_name", None)
    st.session_state.pop("scanned_blocks", None)
    # Slide images are tied to a specific deck -- a cache left over from
    # a previous Analyze would otherwise show the WRONG slide pictures
    # (right slide numbers, stale deck) if the thumbnail URLs happened
    # to still resolve.
    st.session_state.pop("slide_thumbnails", None)
    st.session_state.pop("presentation_id", None)
    st.session_state["slide_browser_page"] = 0
    # "Does this need its own Level?" answers are keyed by block label +
    # title (see render_anchor_resolution) -- a NEW Analyze should start
    # every one of those questions unanswered again, rather than silently
    # reusing a leftover answer from a previous deck/lesson analyzed
    # earlier in the same running session that happened to share a
    # block label and title (e.g. two different lessons both having an
    # "Activity" section titled "Plan").
    for key in [k for k in st.session_state if k.startswith("anchor_choice::")]:
        st.session_state.pop(key, None)
    st.session_state["mode"] = mode
    st.session_state["expected_days"] = expected_days
    if slides_url:
        st.session_state["presentation_id"] = extract_drive_file_id(slides_url)
    try:
        if pptx_override:
            pptx_path = pptx_override
        else:
            with st.spinner("Downloading deck from Google Drive..."):
                pptx_path = "downloaded_deck.pptx"
                download_pptx_via_drive_api(slides_url, CREDENTIALS_PATH, TOKEN_PATH, pptx_path)

        with st.spinner("Extracting speaker notes..."):
            if mode == "Multi-Day Project":
                plan, orphaned_slides, detected_days = extract_project_plan(pptx_path)
                st.session_state["detected_days"] = detected_days
            else:
                plan, orphaned_slides = extract_activity_plan(pptx_path)

        with st.spinner("Reading the lesson's current structure from Levelbuilder..."):
            scanned_blocks = scan_levelbuilder_live(edit_url, SESSION_FILE)
        scanned_by_name = {b["name"].strip(): b["activitySections"] for b in scanned_blocks}

        st.session_state["plan"] = plan
        st.session_state["scanned_by_name"] = scanned_by_name
        st.session_state["scanned_blocks"] = scanned_blocks
        st.session_state["working_edit_url"] = edit_url
        if orphaned_slides:
            st.warning(
                f"Slide(s) {orphaned_slides} had notes before any section "
                f"marker and were skipped -- check the deck."
            )
    except RuntimeError as e:
        st.session_state["needs_login"] = True
        st.session_state["login_error"] = str(e)
    except Exception as e:
        st.error(f"Something went wrong: {e}")
        st.code(traceback.format_exc())


# ----------------------------------------------------------------------
# Login flow (persisted across reruns via session_state, not tied to
# the Analyze button's own transient click state -- a button's click
# result only lasts for the single rerun it triggers, so a login button
# nested under "if analyze_clicked" would never actually fire on its
# own click).
# ----------------------------------------------------------------------

if st.session_state.get("needs_login"):
    st.error(st.session_state.get("login_error", "Your Levelbuilder session may have expired."))
    st.info("Click below to log in -- a browser window will open. Log in there; this page will pick it up automatically once you reach the lesson page, so there's no second button to click.")
    if st.button("Log in to Levelbuilder"):
        try:
            with st.spinner("Browser open -- log in there now. Waiting for you to reach the lesson page (up to 5 minutes)..."):
                login_and_save_session(edit_url, SESSION_FILE)
            st.session_state["needs_login"] = False
            st.success("Session saved. Click Analyze again.")
        except TimeoutError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Something went wrong during login: {e}")
            st.code(traceback.format_exc())


# ----------------------------------------------------------------------
# Step 2: resolve any title/anchor mismatches, then review the plan
# ----------------------------------------------------------------------

if "plan" in st.session_state:
    plan = st.session_state["plan"]
    scanned_by_name = st.session_state["scanned_by_name"]
    active_mode = st.session_state.get("mode", "Single Lesson")

    st.header("Review")

    if active_mode == "Multi-Day Project":
        detected_days = st.session_state.get("detected_days")
        expected_days_val = st.session_state.get("expected_days")
        st.caption(f"Detected {detected_days} day(s) in the deck.")

        # Same "titled sub-activity with no Levels planned" check as
        # single-lesson mode, just run once per (Day, Warm Up/Activity/
        # Wrap Up) sub-block instead of once per fixed block name --
        # confirmed necessary: a day can have more than one titled,
        # self-contained chunk (its own Goal + Transition) even though
        # only one of them has a real Level attached in Levelbuilder
        # (e.g. "Ask Good Questions" vs. "Plan Your Program", or "Peer
        # Review" sending students back to an earlier day's Level vs.
        # "Finish Your Wrapped Program").
        no_anchor_titles = set()
        mismatch_count = 0
        for day_key, sub_name, deck_sections, existing_sections in iter_day_sub_blocks(
            plan, st.session_state["scanned_blocks"]
        ):
            total_anchor_eligible = sum(1 for s in deck_sections if s.get("anchor_eligible"))
            candidate_titles = [s["title"] for s in deck_sections if s.get("anchor_eligible") and s["title"]]
            num_anchors = sum(1 for s in existing_sections if is_real_anchor(s))

            if total_anchor_eligible != num_anchors:
                block_label = f"{day_key} - {sub_name}"
                no_anchor, still_unanswered = render_anchor_resolution(block_label, candidate_titles)
                no_anchor_titles |= no_anchor
                mismatch_count += still_unanswered

        no_anchor_titles = frozenset(no_anchor_titles)

        plan_markdown, all_fill_ops, all_create_groups, project_warnings = build_project_plan(
            plan, st.session_state["scanned_blocks"], expected_days=expected_days_val, no_anchor_titles=no_anchor_titles
        )
        for w in project_warnings:
            st.warning(w)

        preload_plan_thumbnails(all_fill_ops, all_create_groups)
        render_review_summary(all_fill_ops, all_create_groups, plan_markdown, mismatch_count)

    else:
        # For each block, find titled sub-activities that WOULD be anchor-
        # matched by default (anchor_eligible), and compare the count to how
        # many real anchors (sections with Levels) exist in Levelbuilder. A
        # mismatch means some titled sub-activity has no Levels planned --
        # but WHICH one is a judgment call, not something to infer from
        # counts alone.
        no_anchor_titles = set()
        mismatch_count = 0
        for block_name in SECTION_ORDER:
            if block_name not in scanned_by_name:
                st.error(f"'{block_name}' wasn't found in Levelbuilder at all -- check the edit page.")
                continue

            # The REAL count to compare against num_anchors must include
            # every anchor-eligible chunk, even ones that transferred
            # eligibility to a title-less Transition box further down (e.g.
            # "Pixelation Widget: Color" transfers to its own Transition
            # box, which has no title of its own). candidate_titles is a
            # NARROWER subset -- only the ones with an actual title text,
            # since that's all a checkbox can meaningfully label. Using the
            # narrower count for the mismatch check was wrong: it could show
            # a mismatch (and checkboxes) even when the real matching
            # already lines up perfectly, risking an unnecessary and
            # actively WRONG checkbox click on an otherwise-correct plan.
            total_anchor_eligible = sum(1 for s in plan[block_name] if s.get("anchor_eligible"))
            candidate_titles = [s["title"] for s in plan[block_name] if s.get("anchor_eligible") and s["title"]]
            num_anchors = sum(1 for s in scanned_by_name[block_name] if is_real_anchor(s))

            if total_anchor_eligible != num_anchors:
                no_anchor, still_unanswered = render_anchor_resolution(block_name, candidate_titles)
                no_anchor_titles |= no_anchor
                mismatch_count += still_unanswered

        no_anchor_titles = frozenset(no_anchor_titles)

        # Recompute live (cheap, pure logic) on every rerun, so checkbox
        # changes update the plan immediately without re-fetching anything.
        all_fill_ops = []
        all_create_groups = []
        plan_markdown = []
        for block_name in SECTION_ORDER:
            if block_name not in scanned_by_name:
                continue
            all_fill_ops.extend(
                compute_fill_operations(block_name, plan[block_name], scanned_by_name[block_name], no_anchor_titles)
            )
            all_create_groups.extend(
                compute_create_operations(block_name, plan[block_name], scanned_by_name[block_name], no_anchor_titles)
            )
            build_plan_for_block(block_name, plan[block_name], scanned_by_name[block_name], plan_markdown, no_anchor_titles)

        preload_plan_thumbnails(all_fill_ops, all_create_groups)
        render_review_summary(all_fill_ops, all_create_groups, plan_markdown, mismatch_count)

    with st.expander("🖼️ Review each item with its slide image"):
        render_plan_items_with_images(all_fill_ops, all_create_groups)

    st.divider()
    approve_clicked = st.button("Approve and write to Levelbuilder", type="primary")

    with st.expander("🖼️ Browse all slides"):
        render_slide_browser(plan, active_mode)

    if approve_clicked:
        log_area = st.empty()
        log_lines = []

        def log(message):
            log_lines.append(message)
            log_area.code("\n".join(log_lines))

        try:
            with st.spinner("Opening browser and writing to Levelbuilder -- a window will open and stay open for your review..."):
                from playwright.sync_api import sync_playwright

                playwright_instance = sync_playwright().start()
                browser = playwright_instance.firefox.launch(headless=False)
                context = browser.new_context(storage_state=SESSION_FILE)
                page = context.new_page()
                page.goto(st.session_state["working_edit_url"])
                page.wait_for_load_state("networkidle")
                # "networkidle" doesn't guarantee React has actually
                # finished rendering the lesson's sections -- wait for
                # the real UI marker (the first +Activity Section
                # button) to genuinely be there before touching anything.
                page.wait_for_selector("button.add-activity-section", timeout=30000)

                log(f"{len(all_fill_ops)} existing section(s) to fill...")
                run_fill_existing(page, all_fill_ops, log=log)

                log(f"\n{len(all_create_groups)} group(s) of new sections to create...")
                run_create_new(page, all_create_groups, log=log)

                log(
                    "\nAll done. NOTHING HAS BEEN SAVED. Review the lesson in "
                    "the browser window and click Levelbuilder's own 'Save "
                    "and Keep Editing' or 'Save and Close' yourself."
                )

                # Deliberately NOT calling browser.close() or
                # playwright_instance.stop() -- the browser stays open for
                # manual review and save. Stashing references in
                # session_state too (even though nothing calls methods on
                # them again) so Python's garbage collector doesn't tear
                # down the connection the moment this function returns and
                # these local variables go out of scope -- that would
                # close the browser right when we most want it to stay
                # open.
                st.session_state["_write_playwright"] = playwright_instance
                st.session_state["_write_browser"] = browser
                st.session_state["_write_context"] = context
                st.session_state["_write_page"] = page

            st.success("Write complete. Switch to the browser window to review and save.")
        except Exception as e:
            st.error(f"Something went wrong partway through: {e}")
            st.code(traceback.format_exc())
            st.info(
                "If a browser window is open, check it directly -- some "
                "sections may already be filled. Nothing is saved until you "
                "click Save yourself, so it's safe to review before deciding "
                "what to do next."
            )
