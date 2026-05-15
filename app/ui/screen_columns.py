"""
screen_columns.py — Screen 2: Column Matching Confirmation
------------------------------------------------------------
Shows the proposed column matches between V1 and V2, lets the analyst:
  - Confirm auto-matched pairs (all auto-confirmed by default)
  - Manually pair unmatched columns via dropdowns
  - Select the site ID column
  - Set the materiality threshold
  - Review template warnings (if a template was loaded)
  - Review structural site changes (added/removed sites)

Returns a dict on Confirm:
{
    "id_column":             str,       # V1 column name
    "column_mapping":        list[dict],
    "materiality_threshold": float,
}
Returns None if the user goes back or closes.
"""

from __future__ import annotations

import FreeSimpleGUI as sg

from column_matcher import ColumnMatchResult, ColumnMatch, apply_manual_pair
from ui.theme import (
    apply_theme, BG, BG_CARD, TEXT, TEXT_DIM, ACCENT,
    FONT_MAIN, FONT_BOLD, FONT_LARGE, FONT_SMALL, BORDER,
    RED, YELLOW, GREEN, WHITE,
    section_header, label, note, divider,
    primary_button, secondary_button, error_popup,
)


# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

def _match_row_colour(status: str) -> str:
    if status in ("matched", "manual"):
        return GREEN
    if status == "v1_only":
        return RED
    if status == "v2_only":
        return YELLOW
    return TEXT


def _match_status_label(m: ColumnMatch) -> str:
    if m.status == "matched":
        return "✅ matched"
    if m.status == "manual":
        return "🔗 manual"
    if m.status == "v1_only":
        return "⚠️  deleted in V2"
    if m.status == "v2_only":
        return "⚠️  new in V2"
    return ""


def _matches_table(match_result: ColumnMatchResult) -> list:
    """Build a scrollable table of column matches."""
    rows = []
    for m in match_result.matches:
        colour = _match_row_colour(m.status)
        rows.append([
            sg.Text(m.v1_header or "—", font=FONT_MAIN, text_color=colour,
                    background_color=BG_CARD, size=(30, 1), pad=(4, 2)),
            sg.Text(_match_status_label(m), font=FONT_SMALL,
                    text_color=colour, background_color=BG_CARD,
                    size=(18, 1), pad=(4, 2)),
            sg.Text(m.v2_header or "—", font=FONT_MAIN, text_color=colour,
                    background_color=BG_CARD, size=(30, 1), pad=(4, 2)),
        ])
    return rows


def _manual_pair_row(v1_unmatched: list[str], v2_unmatched: list[str]) -> list:
    """Row for manually pairing an unmatched V1 col with a V2 col."""
    if not v1_unmatched or not v2_unmatched:
        return []
    return [
        [divider()],
        [sg.Text("Manually pair an unmatched column:", font=FONT_BOLD,
                 text_color=ACCENT, background_color=BG)],
        [
            sg.Combo(v1_unmatched, key="-MANUAL_V1-", font=FONT_MAIN,
                     size=(28, 1), readonly=True,
                     background_color=BG_CARD, text_color=TEXT),
            sg.Text("  ↔  ", font=FONT_BOLD, text_color=ACCENT, background_color=BG),
            sg.Combo(v2_unmatched, key="-MANUAL_V2-", font=FONT_MAIN,
                     size=(28, 1), readonly=True,
                     background_color=BG_CARD, text_color=TEXT),
            sg.Button("Pair", key="-PAIR-", font=FONT_SMALL,
                      button_color=(WHITE, ACCENT), border_width=0),
        ],
    ]


def _structural_section(match_result: ColumnMatchResult, df_v1, df_v2, id_col: str) -> list:
    """Show added/removed sites if we can compute them (optional, best-effort)."""
    rows = []
    v1_only_cols = [m.v1_header for m in match_result.v1_only]
    v2_only_cols = [m.v2_header for m in match_result.v2_only]

    if v1_only_cols or v2_only_cols:
        rows.append([section_header("Structural Column Changes")])
        if v1_only_cols:
            rows.append([note(f"  ⚠️  Deleted in V2: {', '.join(v1_only_cols)}")])
        if v2_only_cols:
            rows.append([note(f"  ⚠️  New in V2:     {', '.join(v2_only_cols)}")])

    # Site-level structural — only if id_col is available
    if id_col and id_col in df_v1.columns:
        id_col_v2_candidates = [
            m.v2_header for m in match_result.matches
            if m.v1_header == id_col and m.v2_header
        ]
        id_col_v2 = id_col_v2_candidates[0] if id_col_v2_candidates else id_col
        if id_col_v2 in df_v2.columns:
            ids_v1 = set(df_v1[id_col].astype(str))
            ids_v2 = set(df_v2[id_col_v2].astype(str))
            added = sorted(ids_v2 - ids_v1)
            removed = sorted(ids_v1 - ids_v2)
            if added or removed:
                rows.append([section_header("Structural Site Changes")])
            if added:
                rows.append([note(f"  ⚠️  Sites in V2 not in V1 ({len(added)}): "
                                  + ", ".join(added[:10])
                                  + (" …" if len(added) > 10 else ""))])
            if removed:
                rows.append([note(f"  ⚠️  Sites in V1 not in V2 ({len(removed)}): "
                                  + ", ".join(removed[:10])
                                  + (" …" if len(removed) > 10 else ""))])
    return rows


# ---------------------------------------------------------------------------
# Main screen function
# ---------------------------------------------------------------------------

def show_columns_screen(
    match_result: ColumnMatchResult,
    df_v1,
    df_v2,
    v1_label: str,
    v2_label: str,
    default_threshold: float = 5.0,
) -> dict | None:
    """
    Display Screen 2: Column Matching Confirmation.

    Parameters
    ----------
    match_result : ColumnMatchResult
        Output from column_matcher.match_columns(). May be mutated when
        the analyst manually pairs columns.
    df_v1, df_v2 : pd.DataFrame
        Needed to detect structural site changes.
    v1_label, v2_label : str
        Version labels for display.
    default_threshold : float
        Starting materiality threshold.

    Returns
    -------
    dict | None
    """
    apply_theme()

    all_v1 = [m.v1_header for m in match_result.matches if m.v1_header]
    all_v2 = [m.v2_header for m in match_result.matches if m.v2_header]

    # ID column — pre-select if auto-detected
    id_default = match_result.auto_id_column or ""
    id_choices = all_v1  # Any V1 column can be the ID

    def _build_layout():
        v1_unmatched = [m.v1_header for m in match_result.v1_only]
        v2_unmatched = [m.v2_header for m in match_result.v2_only]
        match_table = _matches_table(match_result)
        structural = _structural_section(match_result, df_v1, df_v2,
                                         values.get("-ID_COL-", id_default)
                                         if _layout_built else id_default)

        layout = [
            [sg.Text("GHG QA Tool", font=("Segoe UI", 18, "bold"),
                     text_color=ACCENT, background_color=BG)],
            [note(f"Comparing  {v1_label}  ↔  {v2_label}")],
            [divider()],

            # ── ID column ────────────────────────────────────────────────
            [section_header("Site Identifier Column")],
            [note("Which column uniquely identifies each site (the join key for all comparisons)?")],
            [
                label("Site ID column:", size=(18, 1)),
                sg.Combo(id_choices, default_value=id_default,
                         key="-ID_COL-", font=FONT_MAIN, size=(30, 1),
                         readonly=True, background_color=BG_CARD, text_color=TEXT),
            ],
            [sg.Text("", background_color=BG)],

            # ── Column matches ───────────────────────────────────────────
            [section_header("Column Matches")],
            [
                sg.Text(f"{'V1 Column':<30}", font=FONT_BOLD, text_color=TEXT,
                        background_color=BG_CARD, pad=(4, 2)),
                sg.Text(f"{'Status':<18}", font=FONT_BOLD, text_color=TEXT,
                        background_color=BG_CARD, pad=(4, 2)),
                sg.Text(f"{'V2 Column':<30}", font=FONT_BOLD, text_color=TEXT,
                        background_color=BG_CARD, pad=(4, 2)),
            ],
            [
                sg.Column(
                    match_table,
                    scrollable=True,
                    vertical_scroll_only=True,
                    size=(700, min(300, max(80, len(match_table) * 26))),
                    background_color=BG_CARD,
                    key="-MATCH_COL-",
                )
            ],
            [sg.Text("", background_color=BG)],

            # ── Manual pairing ───────────────────────────────────────────
            *_manual_pair_row(v1_unmatched, v2_unmatched),
            [sg.Text("", background_color=BG)],

            # ── Structural changes ───────────────────────────────────────
            *structural,
            [sg.Text("", background_color=BG)],

            # ── Template warnings ────────────────────────────────────────
            *(
                [
                    [section_header("Template Warnings")],
                    *[[note(f"  ⚠️  {w}")] for w in match_result.template_warnings[:20]],
                    [sg.Text("", background_color=BG)],
                ]
                if match_result.template_warnings else []
            ),

            # ── Materiality threshold ────────────────────────────────────
            [section_header("Materiality Threshold")],
            [
                label("Flag changes greater than:", size=(24, 1)),
                sg.Input(str(default_threshold), key="-THRESHOLD-",
                         font=FONT_MAIN, size=(6, 1),
                         background_color=BG_CARD, text_color=TEXT),
                label("%  as High Impact 🔴"),
            ],
            [note("  Default is 5% — calibrate this for your project before trusting the results.")],
            [sg.Text("", background_color=BG)],

            [divider()],
            [
                primary_button("Confirm & Run Diff →", "-CONFIRM-"),
                sg.Text("  ", background_color=BG),
                secondary_button("← Back", "-BACK-"),
            ],
            [sg.Text("", key="-STATUS-", font=FONT_SMALL, text_color=TEXT_DIM,
                     background_color=BG, size=(70, 1))],
        ]
        return layout

    # First build needs dummy values dict for structural section
    values = {}
    _layout_built = False
    layout = _build_layout()
    _layout_built = True

    window = sg.Window(
        "GHG QA Tool — Column Matching",
        layout,
        background_color=BG,
        finalize=True,
        resizable=True,
        margins=(20, 20),
    )

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, "-BACK-"):
            window.close()
            return None

        # ── Manual pair ────────────────────────────────────────────────────
        if event == "-PAIR-":
            v1h = values.get("-MANUAL_V1-", "")
            v2h = values.get("-MANUAL_V2-", "")
            if not v1h or not v2h:
                error_popup("Please select both a V1 column and a V2 column to pair.")
                continue
            apply_manual_pair(match_result, v1h, v2h)
            window.close()
            # Rebuild with updated match_result
            _layout_built = False
            layout = _build_layout()
            _layout_built = True
            window = sg.Window(
                "GHG QA Tool — Column Matching",
                layout,
                background_color=BG,
                finalize=True,
                resizable=True,
                margins=(20, 20),
            )
            continue

        # ── Confirm ────────────────────────────────────────────────────────
        if event == "-CONFIRM-":
            id_col = values.get("-ID_COL-", "").strip()
            threshold_str = values.get("-THRESHOLD-", "5").strip()

            errors = []
            if not id_col:
                errors.append("Please select the site identifier column to continue.")
            try:
                threshold = float(threshold_str)
                if threshold < 0 or threshold > 100:
                    errors.append("Materiality threshold must be between 0 and 100.")
            except ValueError:
                errors.append("Materiality threshold must be a number (e.g. 5 or 2.5).")
                threshold = default_threshold

            if errors:
                error_popup("\n".join(errors))
                continue

            window.close()
            return {
                "id_column": id_col,
                "column_mapping": match_result.to_serializable(),
                "materiality_threshold": threshold,
            }
