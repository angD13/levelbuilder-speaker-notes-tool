#!/usr/bin/env python3
"""
build_assignment_plan.py

READ-ONLY -- produces a document for you to review, touches nothing in
Levelbuilder. Combines:
  1. The content plan extracted from the deck's .pptx (one entry per
     slide, same logic as speaker_notes_sync.py --dry-run)
  2. The live structure scanned from Levelbuilder (scan_levelbuilder_lesson.py)

...to decide exactly which existing section should receive which
slide's content, and where brand-new sections need to be created for
leftover content, following the rule confirmed with Ang:

  - Existing ANCHOR sections (Progression Title + Levels attached) get
    matched to planned slides that have their own "Title (N mins)"
    line, BY POSITION (1st anchor <-> 1st titled slide, etc) -- never
    by comparing title text, since those can drift.
  - The slides between two titled slides (or before the first / after
    the last) form a "run" of untitled content. Each such run is
    matched to the corresponding run of existing BLANK sections sitting
    between the same two anchors (or before the first / after the
    last), filling them one slide per existing blank, in order.
  - If a run has MORE slides than existing blanks, the leftover slides
    need NEW sections created (noted here with exactly where).
  - If a run has FEWER slides than existing blanks, the leftover
    existing blanks are simply left alone (confirmed fine by Ang).
  - A mismatch in anchor count vs. titled-slide count is flagged and
    skipped rather than guessed.
"""

import argparse
import json
import sys

from speaker_notes_sync import (
    extract_activity_plan,
    SECTION_ORDER,
    warn_if_unmapped_emoji_remain,
    warn_if_icons_were_auto_fixed,
)


def split_into_runs(slide_sections: list, no_anchor_titles: frozenset = frozenset()) -> tuple:
    """
    Given a block's list of per-slide section dicts (in slide order),
    returns (titled_slides, untitled_runs) where untitled_runs has one
    more entry than titled_slides: untitled_runs[i] is the run of
    untitled slides immediately BEFORE titled_slides[i], and
    untitled_runs[-1] is the trailing run after the last titled slide.

    no_anchor_titles: title strings that should NOT be treated as
    anchor-eligible even though they have a title -- for a sub-activity
    that has no Levels and will never get an anchor section (confirmed
    by Ang, not auto-detected). That section's own title/duration are
    still kept intact on the dict for when it's written into Levelbuilder
    -- it's only excluded from the anchor-matching queue, so it flows in
    as part of the surrounding untitled run instead.

    Anchor eligibility otherwise comes from each section's own
    `anchor_eligible` flag (set by mark_anchor_eligibility in
    speaker_notes_sync.py) rather than raw title presence -- a titled
    box whose Transition lives on a later box is NOT anchor-eligible
    itself; that later box is, even though it has no title of its own.
    """
    titled = []
    untitled_runs = [[]]
    for sec in slide_sections:
        eligible = sec.get("anchor_eligible", bool(sec["title"]))
        if sec["title"] in no_anchor_titles:
            eligible = False
        if eligible:
            titled.append(sec)
            untitled_runs.append([])
        else:
            untitled_runs[-1].append(sec)
    return titled, untitled_runs


def is_real_anchor(sec: dict) -> bool:
    """
    A section only counts as a real anchor if it has BOTH a Progression
    Title AND at least one actual Level attached. Confirmed necessary:
    Levelbuilder can have sections with a Progression Title already
    typed in as a placeholder, with zero Levels attached yet -- treating
    those as real anchors massively overcounts how many anchors exist,
    which scrambles positional matching (two sub-activities ended up
    paired with empty placeholder sections instead of their real,
    further-down anchors).
    """
    return bool(sec.get("progressionName")) and bool(sec.get("scriptLevels"))


def split_existing_into_runs(existing_sections: list) -> tuple:
    """Same idea, but for the live Levelbuilder sections: anchors vs.
    the runs of blank sections between them."""
    anchors = []
    blank_runs = [[]]
    for sec in existing_sections:
        if is_real_anchor(sec):
            anchors.append(sec)
            blank_runs.append([])
        else:
            blank_runs[-1].append(sec)
    return anchors, blank_runs


def title_suffix(sec: dict) -> str:
    """A short ' (Title: X, Y mins)' note for FILL/NEW-section bullets --
    so a title that's excluded from anchor-matching (via --no-anchor-title)
    but still has real Title/Duration values to write is visible in the
    review, not just shown for actual Anchors."""
    if not sec["title"]:
        return ""
    duration_part = f", {sec['duration']} mins" if sec["duration"] else ""
    return f" (Title: {sec['title']!r}{duration_part})"


def format_content_preview(sec: dict, max_len: int = 500) -> str:
    text = sec["body_html"]
    callouts_note = ""
    if sec["callouts"]:
        callouts_note = " + " + ", ".join(c["type"] for c in sec["callouts"])
    preview = text[:max_len] + ("..." if len(text) > max_len else "")
    return preview + callouts_note


def indented_code_block(content: str, indent: str) -> str:
    """
    Wraps content in a markdown fenced code block, applying `indent` to
    EVERY line -- including the fence markers themselves and any
    embedded newline in content. Without this, a content line that
    starts with its own "- " (e.g. a bulleted sub-point) breaks out of
    the surrounding list item in most markdown renderers and gets
    rendered as a real list item instead of staying literal.
    """
    lines = content.split("\n")
    indented_lines = "\n".join(f"{indent}{line}" for line in lines)
    return f"{indent}```\n{indented_lines}\n{indent}```"


def compute_fill_operations(
    block_name: str, slide_sections: list, existing_sections: list, no_anchor_titles: frozenset = frozenset()
) -> list:
    """
    Returns a list of {block, position_in_block, slide} dicts -- one per
    EXISTING section that should be filled with a specific slide's
    content, using the same anchor/blank-run matching logic as the
    human-readable plan. position_in_block is the 0-indexed position of
    that section within this block's existing_sections list (matching
    DOM/document order). Does NOT include "leave empty" or "needs a
    NEW section" cases -- this is only for sections that already exist
    and should be filled.
    """
    titled, untitled_runs = split_into_runs(slide_sections, no_anchor_titles)
    anchors, blank_runs = split_existing_into_runs(existing_sections)
    num_pairs = min(len(titled), len(anchors))

    operations = []
    # Track absolute position within existing_sections as we walk through
    # blank_runs and anchors in their original interleaved order.
    abs_pos = 0
    for i in range(num_pairs + 1):
        existing_run = blank_runs[i] if i < len(blank_runs) else []
        plan_run = untitled_runs[i] if i < len(untitled_runs) else []
        for j, existing_sec in enumerate(existing_run):
            if j < len(plan_run):
                operations.append(
                    {"block": block_name, "position_in_block": abs_pos,
                     "section_id": existing_sec["id"], "slide": plan_run[j]}
                )
            abs_pos += 1
        if i < num_pairs:
            operations.append(
                {"block": block_name, "position_in_block": abs_pos,
                 "section_id": anchors[i]["id"], "slide": titled[i]}
            )
            abs_pos += 1
    return operations


def compute_create_operations(
    block_name: str, slide_sections: list, existing_sections: list, no_anchor_titles: frozenset = frozenset(),
    leading_anchor_id=None,
) -> list:
    """
    Returns a list of groups of leftover slides that need brand-new
    sections created for them, one group per contiguous gap. Each group:
    {block, insert_after_section_id, insert_after_anchor_id,
     insert_at_block_start, slides}
    Exactly one of insert_after_section_id / insert_after_anchor_id /
    insert_at_block_start will be meaningful, telling the write step
    where in the page to start inserting.

    leading_anchor_id: used ONLY for the very first run (i == 0) when
    there's no existing blank section to insert after -- e.g. a
    multi-day project's "Warm Up" divider section, which should never
    receive content itself (confirmed by Ang) but DOES give us a real
    id to insert new sections after, instead of falling back to the
    unsupported insert_at_block_start case. Single-lesson callers never
    pass this (there's no such divider id available at a top-level
    block's own start), so behavior there is unchanged.
    """
    titled, untitled_runs = split_into_runs(slide_sections, no_anchor_titles)
    anchors, blank_runs = split_existing_into_runs(existing_sections)
    num_pairs = min(len(titled), len(anchors))

    create_ops = []
    for i in range(num_pairs + 1):
        existing_run = blank_runs[i] if i < len(blank_runs) else []
        plan_run = untitled_runs[i] if i < len(untitled_runs) else []
        if len(plan_run) > len(existing_run):
            leftover = plan_run[len(existing_run):]
            insert_after_section_id = existing_run[-1]["id"] if existing_run else (
                leading_anchor_id if (i == 0 and leading_anchor_id is not None) else None
            )
            create_ops.append({
                "block": block_name,
                "insert_after_section_id": insert_after_section_id,
                "insert_after_anchor_id": anchors[i - 1]["id"] if (not existing_run and i > 0) else None,
                "insert_at_block_start": (not existing_run and i == 0 and leading_anchor_id is None),
                "slides": leftover,
            })
    return create_ops


def build_plan_for_block(
    block_name: str, slide_sections: list, existing_sections: list, out: list, no_anchor_titles: frozenset = frozenset(),
    leading_anchor_id=None,
) -> None:
    out.append(f"\n# {block_name}\n")
    if no_anchor_titles:
        matching_titles = {s["title"] for s in slide_sections if s["title"]} & no_anchor_titles
        if matching_titles:
            out.append(
                f"> ℹ️ Treating {sorted(matching_titles)} as having no anchor "
                f"(confirmed -- no Levels planned for it), so its content "
                f"flows in as ordinary narrative rather than needing a "
                f"matching anchor section.\n"
            )

    titled, untitled_runs = split_into_runs(slide_sections, no_anchor_titles)
    anchors, blank_runs = split_existing_into_runs(existing_sections)

    if len(titled) != len(anchors):
        out.append(
            f"> ⚠️ MISMATCH: deck has {len(titled)} titled slide(s) in this "
            f"block, but Levelbuilder has {len(anchors)} anchor section(s) "
            f"(with Levels attached). Matching only the first "
            f"{min(len(titled), len(anchors))} pair(s) by position; the "
            f"rest are flagged below and skipped -- check manually.\n"
        )

    num_pairs = min(len(titled), len(anchors))

    # Interleave: blank_run[0], anchor[0], blank_run[1], anchor[1], ...
    for i in range(num_pairs + 1):
        # --- the blank run before this anchor (or trailing, if i == num_pairs) ---
        plan_run = untitled_runs[i] if i < len(untitled_runs) else []
        existing_run = blank_runs[i] if i < len(blank_runs) else []

        out.append(f"## Blank-section run {i + 1} of {num_pairs + 1}")
        out.append(f"({len(plan_run)} planned slide(s), {len(existing_run)} existing blank section(s))\n")

        for j, existing_sec in enumerate(existing_run):
            if j < len(plan_run):
                sec = plan_run[j]
                out.append(f"- Existing blank section #{j + 1} (id {existing_sec['id']}) -> FILL with slide {sec['slide']}{title_suffix(sec)}:")
                out.append(indented_code_block(format_content_preview(sec), "  "))
            else:
                out.append(f"- Existing blank section #{j + 1} (id {existing_sec['id']}) -> leave EMPTY (no planned content left for this run)")

        if len(plan_run) > len(existing_run):
            leftover = plan_run[len(existing_run):]
            if existing_run:
                insert_after_desc = f"existing blank section id {existing_run[-1]['id']}"
            elif leading_anchor_id is not None:
                insert_after_desc = f"the divider section id {leading_anchor_id}"
            else:
                insert_after_desc = "(none -- insert at start of this gap)"
            out.append(
                f"- {len(leftover)} leftover slide(s) need NEW sections "
                f"created here (after {insert_after_desc}):"
            )
            for sec in leftover:
                out.append(f"  - NEW section for slide {sec['slide']}{title_suffix(sec)}:")
                out.append(indented_code_block(format_content_preview(sec), "    "))
        out.append("")

        # --- the anchor itself ---
        if i < num_pairs:
            anchor = anchors[i]
            sec = titled[i]
            title_note = f"\"{sec['title']}\", {sec['duration']} mins" if sec["title"] else "no Title/Duration -- carried by an earlier box"
            out.append(
                f"## Anchor: existing section id {anchor['id']} "
                f"(Progression Title: {anchor['progressionName']!r}, "
                f"{len(anchor.get('scriptLevels', []))} level(s)) "
                f"<- slide {sec['slide']} ({title_note})"
            )
            out.append(indented_code_block(format_content_preview(sec), "") + "\n")

    # Flag any unmatched extras on either side
    if len(titled) > len(anchors):
        for sec in titled[len(anchors):]:
            out.append(
                f"> ⚠️ UNMATCHED titled slide {sec['slide']} (\"{sec['title']}\") "
                f"-- no anchor section left to pair it with. Skipped -- "
                f"check whether Levels still need to be added for this "
                f"sub-activity.\n"
            )
    elif len(anchors) > len(titled):
        for anchor in anchors[len(titled):]:
            out.append(
                f"> ⚠️ UNMATCHED anchor section id {anchor['id']} "
                f"(Progression Title: {anchor['progressionName']!r}) -- "
                f"no titled slide left to pair it with. Skipped -- check "
                f"whether this sub-activity's slide is missing its "
                f"\"Title (N mins)\" line.\n"
            )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--scan-json", required=True, help="Output file from scan_levelbuilder_lesson.py")
    parser.add_argument("--out", default="Assignment_Plan.md")
    parser.add_argument(
        "--no-anchor-title",
        action="append",
        default=[],
        help="A titled sub-activity's exact title text that has NO anchor "
        "section (no Levels planned for it, confirmed by you) -- its "
        "content will flow in as ordinary narrative instead of needing a "
        "matching anchor. Repeat this flag for more than one.",
    )
    args = parser.parse_args()
    no_anchor_titles = frozenset(args.no_anchor_title)

    print(f"Extracting content plan from {args.pptx}...")
    plan, orphaned_slides = extract_activity_plan(args.pptx)
    if orphaned_slides:
        print(f"  WARNING: slide(s) {orphaned_slides} had notes before any section marker -- skipped.")

    for name in SECTION_ORDER:
        for sec in plan[name]:
            label = f"{name} / slide {sec['slide']}"
            warn_if_unmapped_emoji_remain(sec["body_html"], label)
            warn_if_icons_were_auto_fixed(sec["icons_fixed"], label)
            if sec.get("duration_placeholder"):
                print(
                    f"  WARNING: {label} has a duration placeholder "
                    f"({sec['duration_placeholder']!r}) instead of a real "
                    f"number -- Duration left blank, needs to be filled in "
                    f"manually.",
                    file=sys.stderr,
                )

    print(f"Loading live structure from {args.scan_json}...")
    with open(args.scan_json, encoding="utf-8") as f:
        scanned_blocks = json.load(f)
    scanned_by_name = {b["name"].strip(): b["activitySections"] for b in scanned_blocks}

    out = []
    for block_name in SECTION_ORDER:
        if block_name not in scanned_by_name:
            out.append(f"\n# {block_name}\n> ⚠️ Not found in the Levelbuilder scan at all -- check the scan file.\n")
            continue
        build_plan_for_block(block_name, plan[block_name], scanned_by_name[block_name], out, no_anchor_titles)

    with open(args.out, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print(f"\nWrote {args.out}. Review it carefully before any write step runs.")


if __name__ == "__main__":
    main()
