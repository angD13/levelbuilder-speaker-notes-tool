#!/usr/bin/env python3
"""
project_assignment_plan.py

Multi-day PROJECT variant of build_assignment_plan.py -- for Connect
Task-style lessons where Levelbuilder groups sections under "Day 1" /
"Day 2" / ... top-level blocks instead of the usual "Warm Up" /
"Activity" / "Wrap Up" blocks, and ALL of Warm Up, Activity, and Wrap
Up for a given day sit as ordinary (stacked) Activity Sections inside
that ONE "Day N" block -- "Warm Up (N mins)" / "Activity (N mins)" /
"Wrap Up (N mins)" exist on both sides only as literal Title-field
text on otherwise-plain divider sections, not as separate containers.

Confirmed with Ang:
  - A new Day starts every time the deck's "Warm Up" marker slide
    reappears (there's no dedicated "Day N" marker in the deck itself
    to key off instead).
  - The Warm Up / Activity / Wrap Up divider sections in Levelbuilder
    are fixed, permanent labels -- they must NEVER receive content
    themselves (not even the very first chunk), because a real
    sub-activity's own title sometimes needs that section's Title
    field, and "Activity" would already be sitting there. All real
    content -- even the first chunk of a sub-part -- gets created as
    brand-new sections placed immediately after the relevant divider.

Everything else (per-slide paragraph/icon/Callout parsing, anchor
matching, fill-vs-create logic) is untouched and reused directly from
speaker_notes_sync.py / build_assignment_plan.py -- a Day's
Warm Up / Activity / Wrap Up each gets handled by the exact same
engine as a normal lesson's three top-level blocks, just called once
per (Day, sub-section) pair instead of once per fixed name.
"""

import re
import sys

from speaker_notes_sync import (
    SECTION_MARKER_RE,
    build_activity_sections,
    mark_anchor_eligibility,
)
from build_assignment_plan import (
    compute_fill_operations,
    compute_create_operations,
    build_plan_for_block,
    is_real_anchor,
)

SUB_SECTIONS = ["Warm Up", "Activity", "Wrap Up"]

DAY_BLOCK_NAME_RE = re.compile(r"^Day\s+(\d+)\s*$")


def extract_project_plan(pptx_path: str) -> tuple:
    """
    Returns (plan, orphaned_slides, detected_day_count) where plan is
    {"Day 1": {"Warm Up": [...], "Activity": [...], "Wrap Up": [...]},
     "Day 2": {...}, ...} -- same per-slide section dicts as
    extract_activity_plan, just bucketed by (day, sub-section) instead
    of just sub-section.

    A new day starts every time the "Warm Up" marker slide reappears
    (confirmed with Ang -- the deck has no dedicated "Day N" marker of
    its own). Content before the first "Warm Up" marker is orphaned,
    same as in extract_activity_plan.
    """
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("Missing dependency: run `pip install python-pptx` first.")

    prs = Presentation(pptx_path)
    plan = {}
    day_count = 0
    current_day_key = None
    current_sub = None
    orphaned_slides = []

    for i, slide in enumerate(prs.slides, start=1):
        if not slide.has_notes_slide:
            continue
        notes_tf = slide.notes_slide.notes_text_frame
        stripped = notes_tf.text.strip()

        marker = SECTION_MARKER_RE.match(stripped)
        if marker:
            name = marker.group(1)
            if name == "Warm Up":
                day_count += 1
                current_day_key = f"Day {day_count}"
                plan[current_day_key] = {name: [] for name in SUB_SECTIONS}
            current_sub = name
            continue

        if not stripped:
            continue

        if current_day_key is None or current_sub is None:
            orphaned_slides.append(i)
            continue

        plan[current_day_key][current_sub].extend(build_activity_sections(i, notes_tf))

    for day_key, subsections in plan.items():
        for sub_name in SUB_SECTIONS:
            subsections[sub_name] = mark_anchor_eligibility(subsections[sub_name])

    return plan, orphaned_slides, day_count


def split_live_day_block(day_sections: list) -> dict:
    """
    Given one "Day N" Levelbuilder block's full, flat activitySections
    list, locates the three divider sections (Title exactly "Warm Up"
    / "Activity" / "Wrap Up", no Levels attached) and slices the list
    into the three sub-ranges that sit strictly between them.

    Returns {"Warm Up": {"sections": [...], "divider_id": ...},
             "Activity": {...}, "Wrap Up": {...}}
    or raises ValueError (with a message meant to surface directly to
    Ang) if a divider is missing, duplicated, or out of the expected
    Warm Up -> Activity -> Wrap Up order.
    """
    divider_positions = {}
    for idx, sec in enumerate(day_sections):
        # NOTE: confirmed against real lesson_scan.json -- a live Activity
        # Section's own Title-field text comes back under the key "name"
        # (not "title", which doesn't exist on live-scanned sections at
        # all; this codebase only ever used "title" for DECK-extracted
        # sections before this function existed).
        title = (sec.get("name") or "").strip()
        if title in SUB_SECTIONS and not is_real_anchor(sec):
            if title in divider_positions:
                raise ValueError(
                    f"Found more than one '{title}' divider section in this "
                    f"Day block -- expected exactly one. Check the Levelbuilder "
                    f"page for a duplicate."
                )
            divider_positions[title] = idx

    missing = [name for name in SUB_SECTIONS if name not in divider_positions]
    if missing:
        raise ValueError(
            f"Couldn't find a {', '.join(missing)!r} divider section (a plain "
            f"section titled exactly 'Warm Up' / 'Activity' / 'Wrap Up', with "
            f"no Levels attached) in this Day block -- it needs to exist "
            f"already, the same way it does for Day 1."
        )

    warm_idx, activity_idx, wrap_idx = (
        divider_positions["Warm Up"], divider_positions["Activity"], divider_positions["Wrap Up"]
    )
    if not (warm_idx < activity_idx < wrap_idx):
        raise ValueError(
            f"The Warm Up / Activity / Wrap Up divider sections in this Day "
            f"block aren't in the expected order -- found them at positions "
            f"{warm_idx}, {activity_idx}, {wrap_idx} (0-indexed). Check the "
            f"Levelbuilder page structure by hand."
        )

    return {
        "Warm Up": {
            "sections": day_sections[warm_idx + 1:activity_idx],
            "divider_id": day_sections[warm_idx]["id"],
        },
        "Activity": {
            "sections": day_sections[activity_idx + 1:wrap_idx],
            "divider_id": day_sections[activity_idx]["id"],
        },
        "Wrap Up": {
            "sections": day_sections[wrap_idx + 1:],
            "divider_id": day_sections[wrap_idx]["id"],
        },
    }


def iter_day_sub_blocks(plan: dict, scanned_blocks: list):
    """
    Yields (day_key, sub_name, deck_sections, existing_sections) for every
    (Day, Warm Up/Activity/Wrap Up) pair present in both the deck plan and
    the live scan -- used by the UI to run the same anchor-count mismatch
    check single-lesson mode already does, per sub-block, BEFORE the real
    fill/create operations are computed (so the resulting no_anchor_titles
    choices can be threaded into build_project_plan). A Day whose live
    structure doesn't have its three dividers in place is silently
    skipped here -- build_project_plan surfaces that as a real warning on
    its own pass, no need to duplicate it.
    """
    scanned_by_name = {b["name"].strip(): b["activitySections"] for b in scanned_blocks}
    detected_days = sorted(
        (k for k in plan if DAY_BLOCK_NAME_RE.match(k)),
        key=lambda k: int(DAY_BLOCK_NAME_RE.match(k).group(1)),
    )
    for day_key in detected_days:
        if day_key not in scanned_by_name:
            continue
        try:
            split_day = split_live_day_block(scanned_by_name[day_key])
        except ValueError:
            continue
        for sub_name in SUB_SECTIONS:
            yield day_key, sub_name, plan[day_key][sub_name], split_day[sub_name]["sections"]


def build_project_plan(
    plan: dict, scanned_blocks: list, expected_days: int = None, no_anchor_titles: frozenset = frozenset()
) -> tuple:
    """
    The project-mode equivalent of the per-block loop in app.py / the
    __main__ of build_assignment_plan.py. Returns:
      (plan_markdown: list[str], all_fill_ops: list, all_create_groups: list,
       warnings: list[str])

    warnings covers day-count mismatches, a Day present in the deck but
    missing from the live scan (or vice versa), and anything
    split_live_day_block itself raised for a given Day (that Day is
    skipped, not guessed at).
    """
    warnings = []
    detected_days = sorted(plan.keys(), key=lambda k: int(DAY_BLOCK_NAME_RE.match(k).group(1)))

    if expected_days is not None and len(detected_days) != expected_days:
        warnings.append(
            f"You said this project has {expected_days} day(s), but the deck's "
            f"speaker notes have {len(detected_days)} 'Warm Up' marker(s) -- "
            f"double check the deck before trusting the plan below."
        )

    scanned_by_name = {b["name"].strip(): b["activitySections"] for b in scanned_blocks}
    live_days = sorted(
        (name for name in scanned_by_name if DAY_BLOCK_NAME_RE.match(name)),
        key=lambda k: int(DAY_BLOCK_NAME_RE.match(k).group(1)),
    )

    deck_only = [d for d in detected_days if d not in scanned_by_name]
    live_only = [d for d in live_days if d not in plan]
    if deck_only:
        warnings.append(f"Day(s) {deck_only} are in the deck but weren't found in Levelbuilder at all.")
    if live_only:
        warnings.append(f"Day(s) {live_only} exist in Levelbuilder but have no matching content in the deck.")

    plan_markdown = []
    all_fill_ops = []
    all_create_groups = []

    for day_key in detected_days:
        if day_key not in scanned_by_name:
            continue
        plan_markdown.append(f"\n# {day_key}\n")
        try:
            split_day = split_live_day_block(scanned_by_name[day_key])
        except ValueError as e:
            warnings.append(f"{day_key}: {e}")
            plan_markdown.append(f"> ⚠️ Skipped -- {e}\n")
            continue

        for sub_name in SUB_SECTIONS:
            block_label = f"{day_key} - {sub_name}"
            sub_sections = split_day[sub_name]["sections"]
            divider_id = split_day[sub_name]["divider_id"]
            deck_sections = plan[day_key][sub_name]

            all_fill_ops.extend(
                compute_fill_operations(block_label, deck_sections, sub_sections, no_anchor_titles)
            )
            all_create_groups.extend(
                compute_create_operations(
                    block_label, deck_sections, sub_sections, no_anchor_titles, leading_anchor_id=divider_id
                )
            )
            build_plan_for_block(
                block_label, deck_sections, sub_sections, plan_markdown, no_anchor_titles, leading_anchor_id=divider_id
            )

    return plan_markdown, all_fill_ops, all_create_groups, warnings
