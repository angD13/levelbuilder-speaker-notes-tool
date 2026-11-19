#!/usr/bin/env python3
"""
speaker_notes_sync.py

Pulls speaker notes out of a Google Slides deck (exported as .pptx), groups
them into Warm Up / Activity / Wrap Up based on the deck's own section-marker
slides, swaps the emoji labels for Levelbuilder's i-class icon markup, and
(optionally) writes the result into the matching fields on a Levelbuilder
lesson edit page via browser automation.

------------------------------------------------------------------------
WHAT'S VALIDATED vs. WHAT'S A GUESS
------------------------------------------------------------------------
Validated against a real deck (U2L7 Logical Operators Day 2) in a live test:
  - pptx export via Google Drive preserves full speaker notes (the native
    Slides "natural language" export does NOT -- it drops notes entirely).
  - Section-marker slides (a slide whose ENTIRE note is just
    "Warm Up (5 mins)" / "Activity (35 mins)" / "Wrap Up (5 mins)") reliably
    mark the boundaries, at least for this deck's convention.

NOT validated -- these are informed placeholders you will need to fix:
  - The CSS selectors in SECTION_FIELD_HINTS below. I have never seen the
    live DOM of levelbuilder-studio.code.org, so these are guesses at the
    *shape* of the problem, not real selectors. See SETUP.md for how to
    find the real ones in ~10 minutes using your browser's inspector.
  - Whether the rich-text field accepts raw HTML (i-class tags) pasted in,
    or needs some other interaction to register the change before Save
    works. The script tries a couple of strategies; expect to adjust.

Run with --dry-run first. It does everything except touch Levelbuilder,
and writes the assembled HTML for each section to local files so you can
eyeball the result before any browser automation runs at all.
------------------------------------------------------------------------
"""

import argparse
import base64
import json
import re
import sys
from pathlib import Path

# ----------------------------------------------------------------------
# 1. EXTRACTION: pull notes out of a local .pptx file
# ----------------------------------------------------------------------

# A section-marker slide's ENTIRE note is just this (e.g. "Warm Up (5 mins)").
# Adjust the label list if a deck uses different section names. The
# duration itself can be any placeholder (e.g. "Wrap Up (x mins)" when
# the timing hasn't been decided yet) -- it isn't used for anything
# downstream of identifying which block this is, so it's never validated
# as a number here.
SECTION_MARKER_RE = re.compile(
    r"^(Warm Up|Activity|Wrap Up)\s*\(\s*\S+\s*min(?:ute)?s?\s*\)\s*$"
)

SECTION_ORDER = ["Warm Up", "Activity", "Wrap Up"]


_DRIVE_FILE_ID_RE = re.compile(r"/d/([a-zA-Z0-9_-]+)")


def extract_drive_file_id(url_or_id: str) -> str:
    """
    Accepts either a raw Drive/Slides file ID or a full Google Slides
    URL (e.g. https://docs.google.com/presentation/d/FILE_ID/edit...)
    and returns just the file ID either way, so you can paste the
    whole URL you copied from your browser's address bar directly.
    """
    match = _DRIVE_FILE_ID_RE.search(url_or_id)
    if match:
        return match.group(1)
    return url_or_id.strip()


def download_pptx_via_drive_api(slides_url_or_id: str, credentials_path: str, token_path: str, out_path: str) -> str:
    """
    Downloads the pptx export directly from Google Drive using the Drive
    API, from just a Slides URL (or a raw file ID) -- no manual export
    step needed. Confirmed working end to end (real browser login, real
    Drive API call) as of this version.

    Requires a Google Cloud OAuth client (credentials.json, Desktop app
    type) and a one-time browser login, which saves a reusable session
    to token_path so future runs don't need you to log in again.
    """
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        import io
    except ImportError:
        sys.exit(
            "Missing dependencies for --slides-url. Run:\n"
            "  pip install google-api-python-client google-auth-oauthlib"
        )

    # Shared with slide_images.py's thumbnail fetching, so both features
    # use ONE Google login / one token.json instead of two. See
    # get_google_credentials's docstring for how it handles a token
    # saved before a scope (like Slides thumbnails) was added here.
    from slide_images import get_google_credentials

    file_id = extract_drive_file_id(slides_url_or_id)
    creds = get_google_credentials(credentials_path, token_path)

    service = build("drive", "v3", credentials=creds)
    request = service.files().export_media(
        fileId=file_id,
        mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    fh = io.FileIO(out_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return out_path


# ----------------------------------------------------------------------
# 2. ICON SWAP
# ----------------------------------------------------------------------

# Maps just the emoji character to its icon tag. The label text itself
# (e.g. "Goal of this activity:") keeps whatever bold/italic formatting
# it already has in the original deck -- we don't need to hardcode or
# re-apply bolding ourselves; see apply_icon_swap.
ICON_MAP = {
    "🎯": '<i class="fa-solid fa-bullseye-arrow"></i>',
    "💬": '<i class="fa fa-comments" aria-hidden="true"></i>',
    "💡": '<i class="fa fa-lightbulb-o" aria-hidden="true"></i>',
    "☑️": '<i class="fa fa-check-square-o" aria-hidden="true"></i>',
    "📄": '<i class="fa fa-file-text-o" aria-hidden="true"></i>',
    "🔁": '<i class="fa fa-refresh" aria-hidden="true"></i>',
    "💻": '<i class="fa fa-desktop" aria-hidden="true"></i>',
    "🎥": '<i class="fa-solid fa-video"></i>',
}

# For the "flag a label with no icon" check below: maps each inline
# label's bold TEXT (not its emoji) to the icon it should have. Built
# from ICON_MAP's own values plus the label text, so there's one source
# of truth for which icon goes with which label.
LABEL_TEXT_TO_ICON = {
    "Goal of this activity": ICON_MAP["🎯"],
    "Discuss": ICON_MAP["💬"],
    "Discussion Goal": ICON_MAP["💡"],
    "Do This": ICON_MAP["☑️"],
    "Distribute": ICON_MAP["📄"],
    "Circulate": ICON_MAP["🔁"],
    "Transition": ICON_MAP["💻"],
    "Video": ICON_MAP["🎥"],
}

# Rough emoji range for the "did we miss one?" sanity check after swapping.
_EMOJI_RE = re.compile(
    "["
    "\U0001F300-\U0001FAFF"
    "\U00002600-\U000027BF"
    "\U0001F1E6-\U0001F1FF"
    "]"
)


_ICON_INSIDE_BOLD_RE = re.compile(r"\*\*(<i[^>]+></i>)(\s*)")

# When the emoji is its OWN complete bold run, separate from the label's
# run (e.g. "**💡** **Discussion Goal:**" instead of one combined run),
# the icon ends up fully wrapped on both sides with nothing else inside
# ("**<icon>**"). Confirmed to cause a real bug: the leftover "**" right
# after the icon doesn't look like "ends with </i>" to the
# missing-icon-detector, so it inserts a SECOND icon. Stripping the
# markers around a fully self-wrapped icon (bolding an empty tag has no
# visible effect anyway) avoids creating that artifact in the first
# place. Must run BEFORE _ICON_INSIDE_BOLD_RE, which handles the
# different case where the icon is only the leading part of a larger
# bold span (e.g. "**<icon> Transition:**").
_ICON_FULLY_WRAPPED_RE = re.compile(r"\*{1,3}(<i[^>]+></i>)\*{1,3}")

# Some decks bold just the label word itself (e.g. a "Circulate" run)
# with the colon as a separate, unbolded run -- producing "**Circulate**:"
# instead of the usual "**Circulate:**". Both are faithful to the deck's
# own formatting, but the mismatch looks inconsistent once rendered, so
# the colon is normalized to sit inside the bold span like every other
# label.
_LABEL_COLON_OUTSIDE_BOLD_RE = re.compile(
    r"\*\*(" + "|".join(re.escape(label) for label in LABEL_TEXT_TO_ICON) + r")\*\*:"
)


def apply_icon_swap(text: str) -> str:
    for emoji, icon_tag in ICON_MAP.items():
        text = text.replace(emoji, icon_tag)
    # If the emoji happened to sit inside the deck's own bold run (e.g.
    # "**<icon> Transition:**"), move the icon tag outside the bold span
    # so it matches Ang's convention ("<icon> **Transition:**") -- an
    # empty <i> tag has no text to visibly bold anyway, so this is a
    # cosmetic normalization, not a formatting change.
    text = _ICON_FULLY_WRAPPED_RE.sub(r"\1", text)
    text = _ICON_INSIDE_BOLD_RE.sub(r"\1\2**", text)
    text = _LABEL_COLON_OUTSIDE_BOLD_RE.sub(r"**\1:**", text)
    return text


def warn_if_unmapped_emoji_remain(text: str, section_name: str) -> None:
    leftovers = _EMOJI_RE.findall(text)
    if leftovers:
        unique = sorted(set(leftovers))
        print(
            f"  WARNING: unmapped emoji {unique} found in '{section_name}' "
            f"after icon swap -- these were left as-is. Add them to "
            f"ICON_MAP if they should become an icon.",
            file=sys.stderr,
        )


def insert_missing_icons(text: str) -> tuple:
    """
    Some decks are inconsistent about including a label's emoji (e.g.
    "Discussion Goal:" with no preceding lightbulb) -- apply_icon_swap
    can't add an icon that was never there, and that failure is silent
    (nothing is "unmapped", the emoji just never existed). This scans
    for each known label's bold text and, wherever it's NOT already
    preceded by its icon tag, inserts the REAL icon tag directly --
    matching exactly what would have happened if the deck's emoji had
    been there in the first place. Returns (new_text, labels_fixed) so
    the caller can warn about it for review purposes, WITHOUT that
    warning itself ending up as visible text in the final content
    (confirmed: a bracketed "[ADD ICON: ...]" note previously got
    written straight into the real lesson, which is not what's wanted).
    """
    labels_fixed = []
    for label, icon in LABEL_TEXT_TO_ICON.items():
        marker = f"**{label}:**"
        start = 0
        while True:
            idx = text.find(marker, start)
            if idx == -1:
                break
            preceding = text[:idx].rstrip()
            if not preceding.endswith("</i>"):
                insertion = f"{icon} "
                text = text[:idx] + insertion + text[idx:]
                idx += len(insertion)
                labels_fixed.append(label)
            start = idx + len(marker)
    return text, labels_fixed


def warn_if_icons_were_auto_fixed(labels_fixed: list, section_name: str) -> None:
    if labels_fixed:
        print(
            f"  WARNING: '{section_name}' had {sorted(set(labels_fixed))} "
            f"missing its icon in the deck -- the icon was inserted "
            f"automatically. Consider adding the matching emoji to the "
            f"deck itself so this doesn't need patching every time.",
            file=sys.stderr,
        )


# ----------------------------------------------------------------------
# 3. HTML ASSEMBLY
# ----------------------------------------------------------------------


def assemble_html(text: str) -> tuple:
    """
    Icon-swaps a block of text and returns (text, labels_fixed) as plain
    text, preserving the original line breaks exactly. Levelbuilder's
    textarea fields take plain text with inline markdown-ish markup (the
    i-class tag, **bold**, *italic*, "- " bullets) directly -- NOT HTML
    <p>/<br> tags. Confirmed against a real field's content by Ang.
    labels_fixed lists any labels that were missing their icon and had
    one inserted automatically -- purely for the caller to warn about,
    not part of the content itself.
    """
    swapped = apply_icon_swap(text)
    fixed, labels_fixed = insert_missing_icons(swapped)
    return fixed.strip(), labels_fixed


# ----------------------------------------------------------------------
# 3b. ACTIVITY-SECTION PLANNING (one per slide, matching Levelbuilder's
#     real structure: Title/Duration/Textarea, plus Callouts and a
#     Remarks checkbox pulled out of specific labeled blocks)
# ----------------------------------------------------------------------

# A slide's first paragraph counts as a Title/Duration line ONLY if the
# entire paragraph is just this pattern -- e.g. "Predict & Run (10 mins)"
# with nothing else in that paragraph. The duration can be a placeholder
# (e.g. "Investigate (x mins)") when it hasn't been decided yet -- caught
# downstream and left blank with a warning, rather than being required
# to be a number here.
TITLE_DURATION_RE = re.compile(
    r"^(?P<title>.{1,80}?)\s*\(\s*(?P<duration>\S+?)\s*min(?:ute)?s?\s*\)\s*$"
)

# Labels whose block becomes a separate Callout object instead of staying
# inline in the main textarea. Confirmed by Ang: Teaching Tip and
# Assessment Opportunity; everything else stays inline.
CALLOUT_LABELS = {"Teaching Tip", "Assessment Opportunity"}


def _is_callout_label(first_plain: str) -> bool:
    """
    Matches a callout-type label even when the deck writes it with a
    trailing colon ("Teaching Tip:" vs "Teaching Tip") -- confirmed to
    vary between decks, and an exact-match-only check silently misses
    these, letting Teaching Tip/Assessment Opportunity content render
    as ordinary inline text instead of becoming a real Callout.
    """
    return first_plain.rstrip(":").strip() in CALLOUT_LABELS


def _callout_label_key(first_plain: str) -> str:
    """The canonical (colon-free) label text, for use as the callout's
    recorded type once _is_callout_label has confirmed a match."""
    return first_plain.rstrip(":").strip()

# Matches a "Say:" label, tolerating the mic emoji that precedes it in
# the deck (e.g. "🎙️ Say:") and bold markers that can land in different
# places depending on how the deck's author bolded it -- sometimes the
# whole "🎙️ Say:" is one bold run ("**🎙️ Say:**"), sometimes just the
# word "Say" is ("🎙️ **Say**:"). Optional asterisk groups around every
# plausible position handle both. The label itself is discarded entirely
# once detected -- the Remarks checkbox is what communicates "this is a
# Say line" in Levelbuilder, not leftover label text.
SAY_LABEL_RE = re.compile(
    r"^(\*{1,3})?🎙️?\s*(\*{1,3})?Say(\*{1,3})?:?\s*(\*{1,3})?\s*"
)


def _run_markdown(text: str, bold: bool, italic: bool) -> str:
    """Wraps a run's text in markdown bold/italic, keeping surrounding
    whitespace OUTSIDE the markers so the result is valid markdown
    (e.g. 'Level 1: ' bold -> '**Level 1:** ', not '**Level 1: **')."""
    if not bold and not italic:
        return text
    stripped = text.strip()
    if not stripped:
        return text
    lead_ws = text[: len(text) - len(text.lstrip())]
    trail_ws = text[len(text.rstrip()):]
    marker = "***" if (bold and italic) else ("**" if bold else "*")
    return f"{lead_ws}{marker}{stripped}{marker}{trail_ws}"


def get_paragraphs(notes_slide) -> list:
    """
    Returns a list of (plain_text, rich_text) tuples, one per paragraph
    in the notes, in original order. plain_text has no markdown (used
    for pattern-matching labels); rich_text preserves the deck's own
    bold/italic formatting as markdown. Blank paragraphs appear as
    ("", "") and are used downstream as block separators.
    """
    result = []
    for p in notes_slide.paragraphs:
        plain_parts = []
        rich_parts = []
        for r in p.runs:
            if not r.text:
                continue
            plain_parts.append(r.text)
            rich_parts.append(
                _run_markdown(r.text, bool(r.font.bold), bool(r.font.italic))
            )
        result.append(("".join(plain_parts).strip(), "".join(rich_parts).strip()))
    return result


def group_into_blocks(paragraphs: list) -> list:
    """
    Groups (plain, rich) paragraph tuples into blocks: consecutive
    non-blank paragraphs with no blank paragraph between them. This is
    what distinguishes "Transition: / Level 1: / Level 2:" (one block,
    no blank lines between them in the deck) from "Circulate: ..." as
    its own separate block (a blank paragraph precedes it).
    """
    blocks = []
    current = []
    for plain, rich in paragraphs:
        if not plain:
            if current:
                blocks.append(current)
                current = []
            continue
        current.append((plain, rich))
    if current:
        blocks.append(current)
    return blocks


def _split_block_for_callouts(block: list) -> tuple:
    """
    Splits a block (list of (plain, rich) paragraph tuples) into
    (ordinary_paragraphs, callouts), finding a Callout label anywhere
    within the block -- not just as its very first paragraph. Everything
    before the first Callout label stays ordinary content; from that
    label onward, paragraphs belong to that Callout, until either
    another Callout label starts a new one or the block ends.
    """
    ordinary = []
    callouts = []  # list of (label, [body paragraphs])
    current = None  # (label, [body paragraphs]) or None
    for para in block:
        plain, _ = para
        if _is_callout_label(plain):
            if current is not None:
                callouts.append(current)
            current = (_callout_label_key(plain), [])
            continue
        if current is not None:
            current[1].append(para)
        else:
            ordinary.append(para)
    if current is not None:
        callouts.append(current)
    return ordinary, callouts


def _render_block_into(block: list, is_say_block: bool, body_lines: list, icons_fixed: list) -> None:
    """Appends one block's rendered lines (first line + bulleted
    follow-ups) into body_lines, stripping the Say label if present.
    Any auto-fixed icon labels are appended to icons_fixed for the
    caller to warn about."""
    first_plain, first_rich = block[0]
    rest = block[1:]
    if is_say_block:
        first_rich = SAY_LABEL_RE.sub("", first_rich, count=1)
    first_line, fixed = assemble_html(first_rich)
    icons_fixed.extend(fixed)
    block_lines = [first_line]
    for _, r in rest:
        line, fixed = assemble_html(r)
        icons_fixed.extend(fixed)
        block_lines.append(f"- {line}")
    body_lines.append("\n".join(block_lines))


def build_activity_sections(slide_num: int, notes_slide) -> list:
    """
    Turns one slide's notes into ONE OR MORE Levelbuilder Activity
    Sections. Normally one slide -> one section, but a "Say:" block
    ALWAYS becomes its own isolated section -- Levelbuilder's Remarks
    checkbox applies to an entire section's textarea, so a Say line
    can't share a textarea with non-Say content without mislabeling all
    of it as a remark (confirmed against a real example). Everything
    else (bold/italic preservation, bullet follow-ups, Callout
    extraction) works the same as before, just at this finer grain.

    Title/Duration (from a "Title (N mins)" first line) are attached
    ONLY to the first resulting section -- a working assumption, since
    a slide that starts directly with a Say block hasn't been seen in
    practice; flag it if that turns out wrong.
    """
    paragraphs = get_paragraphs(notes_slide)
    blocks = group_into_blocks(paragraphs)

    title = None
    duration = None
    duration_placeholder = None
    if blocks and len(blocks[0]) == 1:
        m = TITLE_DURATION_RE.match(blocks[0][0][0])
        if m:
            title = m.group("title").strip()
            duration_text = m.group("duration")
            if duration_text.isdigit():
                duration = duration_text
            else:
                # A placeholder like "x" -- leave Duration blank rather
                # than typing non-numeric text into it; flagged below.
                duration_placeholder = duration_text
            blocks = blocks[1:]

    # Pass 1: split blocks into runs, isolating every Say block into its
    # own single-block run. Callouts never force a split -- they travel
    # with whichever run they appear in.
    runs = []  # list of (run_blocks, is_say_run)
    current_run = []
    for block in blocks:
        first_plain, _ = block[0]
        if _is_callout_label(first_plain):
            current_run.append(block)
            continue
        if SAY_LABEL_RE.match(first_plain):
            if current_run:
                runs.append((current_run, False))
                current_run = []
            runs.append(([block], True))
            continue
        current_run.append(block)
    if current_run:
        runs.append((current_run, False))

    if not runs:
        runs = [([], False)]

    # Pass 2: render each run into its own section dict.
    sections = []
    for idx, (run_blocks, is_say_run) in enumerate(runs):
        callouts = []
        body_lines = []
        icons_fixed = []
        for block in run_blocks:
            # A Callout label can appear ANYWHERE within a block, not
            # just as its very first paragraph -- confirmed: a deck can
            # run "Transition: / Level 1: / Level 2: / Teaching Tip:"
            # with no blank line anywhere in between, which the coarser
            # blank-line grouping treats as one block. Scanning the
            # whole block (not just block[0]) catches this.
            ordinary_paras, found_callouts = _split_block_for_callouts(block)
            for label, body_paras in found_callouts:
                callout_parts = []
                for _, r in body_paras:
                    line, fixed = assemble_html(r)
                    icons_fixed.extend(fixed)
                    callout_parts.append(line)
                callouts.append({"type": label, "text": "\n\n".join(callout_parts)})
            if ordinary_paras:
                _render_block_into(ordinary_paras, is_say_run, body_lines, icons_fixed)

        sections.append({
            "slide": slide_num,
            "title": title if idx == 0 else None,
            "duration": duration if idx == 0 else None,
            "duration_placeholder": duration_placeholder if idx == 0 else None,
            "remarks_checked": is_say_run,
            "callouts": callouts,
            "body_html": "\n\n".join(body_lines).strip(),
            "icons_fixed": icons_fixed,
        })

    return sections


def mark_anchor_eligibility(sections: list) -> list:
    """
    Some decks put a sub-activity's "Goal" on one slide and its
    "Transition" (the one that actually introduces the specific Levels)
    on a LATER slide, rather than together on one slide (confirmed: the
    same deck does both, inconsistently, depending on the sub-activity).
    The box that should be matched against the Levelbuilder anchor (the
    one with Levels attached) is whichever one contains the Transition
    -- not necessarily the one carrying the slide's title line.

    This does NOT merge content together -- each box stays separate, in
    its original chronological order (so Say/Teaching-Tip boxes that sit
    between Goal and Transition keep their correct position). It only
    adds an `anchor_eligible` flag to each section dict: normally the
    same as "has a title", but when a titled box doesn't itself contain
    a Transition and a later box does, eligibility transfers to that
    later box -- the titled box keeps its Title/Duration for display
    (as an ordinary, non-anchor box) but drops out of anchor matching.
    """
    transition_icon = ICON_MAP["💻"]
    video_icon = ICON_MAP["🎥"]
    distribute_icon = ICON_MAP["📄"]
    do_this_icon = ICON_MAP["☑️"]

    def is_skippable(sec: dict) -> bool:
        # Content that's clearly just a short aside still introducing
        # THIS SAME sub-activity -- safe to scan past looking for the
        # real Transition. Deliberately NOT "any untitled content":
        # confirmed that scanning past substantial, unrelated content
        # (a Discuss block, a Discussion Goal) can wander many slides
        # and wrongly attach to a LATER, unrelated anchor's own
        # Transition -- a real bug, not a hypothetical one. Video,
        # Distribute, and plain Do This boxes are all confirmed safe to
        # add: a genuine Discuss/Discussion Goal still blocks the scan
        # before it could wander too far (that's what actually kept
        # U1L8 correct, not Do This/Distribute/Video being blocked).
        if sec["title"]:
            return False
        if sec["remarks_checked"] or not sec["body_html"].strip():
            return True  # a Say box, or a Callout-only box with no real body
        if sec["body_html"].startswith((video_icon, distribute_icon, do_this_icon)):
            return True
        return False

    result = [dict(sec) for sec in sections]
    for sec in result:
        sec["anchor_eligible"] = bool(sec["title"])

    for i, sec in enumerate(result):
        if not sec["title"] or sec["body_html"].startswith(transition_icon):
            continue  # not titled, or already the simple same-box case -- nothing to transfer
        j = i + 1
        match = None
        while j < len(result):
            candidate = result[j]
            if candidate["title"]:
                break  # reached the next titled slide -- no Transition found for this one
            # Check for a real match BEFORE checking skippability: the
            # Transition can be a second or third labeled paragraph
            # inside an otherwise-plain box (e.g. a "Do This" box that
            # ends with its own "Transition:" line), not only a box
            # that starts with the Transition icon as its first label.
            if transition_icon in candidate["body_html"]:
                match = j
                break
            if not is_skippable(candidate):
                break  # genuine blocking content (e.g. Discuss) -- stop here
            j += 1
        if match is not None:
            sec["anchor_eligible"] = False
            result[match]["anchor_eligible"] = True

    return result


def extract_activity_plan(pptx_path: str) -> tuple:
    """
    Returns (plan, orphaned_slides) where plan is
    {"Warm Up": [section_dict, ...], "Activity": [...], "Wrap Up": [...]}
    -- one section_dict per slide, in slide order, matching one
    Levelbuilder Activity Section each.
    """
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("Missing dependency: run `pip install python-pptx` first.")

    prs = Presentation(pptx_path)
    plan = {name: [] for name in SECTION_ORDER}
    current = None
    orphaned_slides = []

    for i, slide in enumerate(prs.slides, start=1):
        if not slide.has_notes_slide:
            continue
        notes_tf = slide.notes_slide.notes_text_frame
        stripped = notes_tf.text.strip()  # plain text, just for marker/empty checks

        marker = SECTION_MARKER_RE.match(stripped)
        if marker:
            current = marker.group(1)
            continue

        if not stripped:
            continue

        if current is None:
            orphaned_slides.append(i)
            continue

        plan[current].extend(build_activity_sections(i, notes_tf))

    for name in SECTION_ORDER:
        plan[name] = mark_anchor_eligibility(plan[name])

    return plan, orphaned_slides


def write_plan_markdown(plan: dict, orphaned_slides: list, out_path: str) -> None:
    lines = []
    if orphaned_slides:
        lines.append(
            f"> WARNING: slide(s) {orphaned_slides} had notes before any "
            f"section marker and were skipped entirely. Check these "
            f"manually.\n"
        )

    for section_name in SECTION_ORDER:
        sections = plan[section_name]
        lines.append(f"# {section_name}  (top-level Activity)\n")
        if not sections:
            lines.append("_(no content found -- check the marker slide exists)_\n")
            continue

        for idx, sec in enumerate(sections, start=1):
            title_display = sec["title"] or "(blank -- no \"Title (N mins)\" line found on this slide)"
            if sec.get("duration_placeholder"):
                duration_display = (
                    f"(blank -- deck has placeholder {sec['duration_placeholder']!r}, "
                    f"needs to be filled in manually)"
                )
            else:
                duration_display = sec["duration"] or "(blank)"
            lines.append(f"## Activity Section {idx}  _(from slide {sec['slide']})_")
            lines.append(f"- **Title field:** {title_display}")
            lines.append(f"- **Duration field:** {duration_display}")
            lines.append(
                f"- **Remarks checkbox:** "
                f"{'CHECKED' if sec['remarks_checked'] else 'unchecked'}"
            )
            if sec["callouts"]:
                for c in sec["callouts"]:
                    preview = c["text"][:100] + ("..." if len(c["text"]) > 100 else "")
                    lines.append(f"- **Callout to add ({c['type']}):** \"{preview}\"")
            else:
                lines.append("- **Callouts:** none")
            lines.append("")
            lines.append("**Textarea content for this section:**")
            lines.append("```")
            lines.append(sec["body_html"] if sec["body_html"] else "(empty)")
            lines.append("```")
            lines.append("")

    Path(out_path).write_text("\n".join(lines), encoding="utf-8")


# ----------------------------------------------------------------------
# 4. LEVELBUILDER WRITE (Playwright) -- UNVERIFIED, see SETUP.md
# ----------------------------------------------------------------------

# TODO: Replace these with real selectors found via your browser's
# inspector (see SETUP.md, step "Finding the real selectors"). What's
# here is a best-guess starting point, not a tested value.
SECTION_FIELD_HINTS = {
    "Warm Up": {
        "heading_text": "Warm Up",
        # e.g. "#lesson_warm_up" or "textarea[name='lesson[warm_up]']"
        "field_selector": None,
    },
    "Activity": {
        "heading_text": "Activity",
        "field_selector": None,
    },
    "Wrap Up": {
        "heading_text": "Wrap Up",
        "field_selector": None,
    },
}


def login_once_and_save_session(edit_url: str, storage_state_path: str, browser_name: str) -> None:
    """
    One-time step: opens a real browser window, lets you log into
    Levelbuilder by hand, then saves the session so future runs don't
    need you to log in again. Re-run this if your session ever expires.
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser_type = getattr(p, browser_name)
        browser = browser_type.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(edit_url)
        input(
            "\nA browser window opened. Log into Levelbuilder, navigate to "
            "the lesson edit page if it redirected you, then come back "
            "here and press Enter to save the session...\n"
        )
        context.storage_state(path=storage_state_path)
        browser.close()
    print(f"Session saved to {storage_state_path}")


def write_sections_to_levelbuilder(
    section_html: dict,
    edit_url: str,
    storage_state_path: str,
    browser_name: str,
    headless: bool,
) -> None:
    from playwright.sync_api import sync_playwright

    missing_selectors = [
        name
        for name, hint in SECTION_FIELD_HINTS.items()
        if not hint["field_selector"]
    ]
    if missing_selectors:
        sys.exit(
            f"SECTION_FIELD_HINTS is missing field_selector for: "
            f"{missing_selectors}. Fill these in first -- see SETUP.md, "
            f"'Finding the real selectors'. Nothing was written."
        )

    with sync_playwright() as p:
        browser_type = getattr(p, browser_name)
        browser = browser_type.launch(headless=headless)
        context = browser.new_context(storage_state=storage_state_path)
        page = context.new_page()
        page.goto(edit_url)
        page.wait_for_load_state("networkidle")

        for section_name, html in section_html.items():
            selector = SECTION_FIELD_HINTS[section_name]["field_selector"]
            locator = page.locator(selector)
            locator.wait_for(state="visible", timeout=10_000)

            # Strategy A: contenteditable rich-text div -- set innerHTML
            # directly and fire an input event so the editor's internal
            # state (and anything wired to Save) picks up the change.
            try:
                locator.evaluate(
                    """(el, html) => {
                        el.innerHTML = html;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                    }""",
                    html,
                )
                print(f"  Wrote '{section_name}' via innerHTML + input event.")
            except Exception as e:
                print(
                    f"  Strategy A failed for '{section_name}' ({e}). "
                    f"This field may be a plain textarea or an iframe-based "
                    f"editor instead -- see SETUP.md troubleshooting.",
                    file=sys.stderr,
                )

        print(
            "\nDone writing fields. NOTHING HAS BEEN SAVED YET. "
            "Review the lesson in the browser and click 'Save and Keep "
            "Editing' or 'Save and Close' yourself -- this script does "
            "not click Save automatically, on purpose."
        )
        if not headless:
            input("Press Enter to close the browser...")
        browser.close()


# ----------------------------------------------------------------------
# 5. CLI
# ----------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=False)
    source.add_argument("--pptx", help="Path to a locally-exported .pptx file")
    source.add_argument(
        "--slides-url",
        help="The Google Slides deck's URL (or just its file ID) -- pulls "
        "the deck directly via the Drive API, no manual export needed. "
        "Requires --credentials.",
    )
    parser.add_argument(
        "--credentials", default="credentials.json",
        help="OAuth client secrets JSON (only for --slides-url)",
    )
    parser.add_argument(
        "--token", default="token.json",
        help="Where the saved Google login session is stored (only for --slides-url)",
    )
    parser.add_argument(
        "--edit-url",
        help="Levelbuilder lesson edit page URL (required unless --dry-run)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Extract, group, and icon-swap only -- write HTML to local "
        "files instead of touching Levelbuilder. Do this first.",
    )
    parser.add_argument(
        "--login",
        action="store_true",
        help="Run the one-time interactive login step and exit.",
    )
    parser.add_argument(
        "--browser",
        choices=["firefox", "chromium", "webkit"],
        default="firefox",
        help="Which browser Playwright should drive (default: firefox)",
    )
    parser.add_argument(
        "--session-file",
        default="levelbuilder_session.json",
        help="Where the saved login session is stored",
    )
    parser.add_argument(
        "--out-dir",
        default="./notes_preview",
        help="Where --dry-run writes its preview files",
    )
    args = parser.parse_args()

    if args.login:
        if not args.edit_url:
            sys.exit("--login requires --edit-url so it knows where to go.")
        login_once_and_save_session(args.edit_url, args.session_file, args.browser)
        return

    if not args.pptx and not args.slides_url:
        sys.exit("Either --pptx or --slides-url is required (unless using --login).")

    # --- Step 1: get the pptx ---
    if args.slides_url:
        pptx_path = "downloaded_deck.pptx"
        print("Downloading deck from Google Drive...")
        download_pptx_via_drive_api(args.slides_url, args.credentials, args.token, pptx_path)
    else:
        pptx_path = args.pptx

    # --- Step 2: extract into the real Activity Section structure ---
    print(f"Extracting notes from {pptx_path}...")
    plan, orphaned_slides = extract_activity_plan(pptx_path)

    if orphaned_slides:
        print(
            f"  WARNING: slide(s) {orphaned_slides} had notes before any "
            f"section marker -- skipped. See the plan file for detail.",
            file=sys.stderr,
        )
    for name in SECTION_ORDER:
        if not plan[name]:
            print(f"  WARNING: no content found for section '{name}'.", file=sys.stderr)
        for sec in plan[name]:
            warn_if_unmapped_emoji_remain(sec["body_html"], f"{name} / slide {sec['slide']}")
            warn_if_icons_were_auto_fixed(sec["icons_fixed"], f"{name} / slide {sec['slide']}")
            if sec.get("duration_placeholder"):
                print(
                    f"  WARNING: {name} / slide {sec['slide']} has a "
                    f"duration placeholder ({sec['duration_placeholder']!r}) "
                    f"instead of a real number -- Duration left blank, "
                    f"needs to be filled in manually.",
                    file=sys.stderr,
                )

    if args.dry_run:
        out_dir = Path(args.out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        plan_path = out_dir / "Activity_Plan.md"
        write_plan_markdown(plan, orphaned_slides, str(plan_path))
        total_sections = sum(len(v) for v in plan.values())
        print(
            f"\nWrote {plan_path} -- {total_sections} Activity Sections "
            f"planned across Warm Up / Activity / Wrap Up."
        )
        print(
            "Open that file and compare it, section by section, against "
            "what's currently in Levelbuilder for this lesson. Do NOT "
            "proceed to a real write until this plan looks right to you."
        )
        return

    sys.exit(
        "The automated Levelbuilder write step is being rebuilt to match "
        "the real Activity Section structure (Title/Duration/Textarea "
        "creation + Callout modals) and isn't ready yet. Run with "
        "--dry-run and review Activity_Plan.md for now."
    )

    # --- Step 4: write to Levelbuilder ---
    if not args.edit_url:
        sys.exit("--edit-url is required unless --dry-run")
    if not Path(args.session_file).exists():
        sys.exit(
            f"No saved session at {args.session_file}. Run with --login "
            f"--edit-url <url> first."
        )

    write_sections_to_levelbuilder(
        section_html,
        args.edit_url,
        args.session_file,
        args.browser,
        headless=False,  # headed by default so you can watch and intervene
    )


if __name__ == "__main__":
    main()
