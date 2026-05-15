"""
screen_project.py — Screen 1: Project Setup
--------------------------------------------
Analyst enters their name, names the project, selects two .xlsx files,
picks the sheet to diff in each, and optionally loads a template.

Returns a dict on success:
{
    "analyst":      str,
    "project":      str,
    "v1_path":      str,
    "v1_label":     str,
    "v1_sheet":     str,
    "v2_path":      str,
    "v2_label":     str,
    "v2_sheet":     str,
    "template_path": str | None,
}
Returns None if the user closes the window.
"""

from __future__ import annotations

import json
from pathlib import Path

import FreeSimpleGUI as sg

from ui.theme import (
    apply_theme, BG, BG_CARD, TEXT, TEXT_DIM, ACCENT, FONT_MAIN, FONT_BOLD,
    FONT_LARGE, FONT_SMALL, BORDER, GREEN, WHITE,
    section_header, label, note, divider, primary_button, secondary_button,
    error_popup,
)
from diff_engine import list_sheets

CONFIG_FILE = Path(__file__).parent.parent.parent / "ghg_tool_config.json"


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

def _load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"analyst_name": "", "recent_projects": [], "default_materiality_threshold": 5.0}


def _save_config(config: dict) -> None:
    try:
        CONFIG_FILE.write_text(json.dumps(config, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# File-loading helper
# ---------------------------------------------------------------------------

def _load_file_sheets(path: str, slot: str, window: sg.Window) -> list[str]:
    """
    Try to load sheet names from an Excel file.
    On error, shows an error popup and returns [].
    Updates the window's sheet dropdown for the given slot.
    """
    if not path or not Path(path).exists():
        return []
    try:
        sheets = list_sheets(path)
    except PermissionError as exc:
        error_popup(str(exc))
        return []
    except Exception as exc:
        error_popup(
            f"Could not read '{Path(path).name}'.\n\n"
            f"Make sure the file is a valid .xlsx file and is not open in Excel.\n\n"
            f"Detail: {exc}"
        )
        return []

    if not sheets:
        error_popup(f"'{Path(path).name}' has no sheets.")
        return []

    window[f"-{slot}_SHEET-"].update(values=sheets, value=sheets[0])
    window[f"-{slot}_SHEET_ROW-"].update(visible=True)
    return sheets


# ---------------------------------------------------------------------------
# Layout builder
# ---------------------------------------------------------------------------

def _file_row(slot: str, label_text: str) -> list:
    """Build the layout rows for one file slot (V1 or V2)."""
    return [
        [
            sg.Text(label_text, font=FONT_BOLD, text_color=ACCENT,
                    background_color=BG, size=(12, 1)),
            sg.Input(key=f"-{slot}_PATH-", font=FONT_MAIN, size=(45, 1),
                     background_color=BG_CARD, text_color=TEXT,
                     readonly=False, enable_events=True),
            sg.FileBrowse("Browse…", target=f"-{slot}_PATH-",
                          file_types=(("Excel Files", "*.xlsx"),),
                          font=FONT_SMALL,
                          button_color=(TEXT, BORDER)),
        ],
        [
            sg.Text("Label:", font=FONT_MAIN, text_color=TEXT_DIM,
                    background_color=BG, size=(12, 1)),
            sg.Input(key=f"-{slot}_LABEL-", font=FONT_MAIN, size=(25, 1),
                     default_text="Draft v1" if slot == "V1" else "Draft v2",
                     background_color=BG_CARD, text_color=TEXT),
        ],
        [
            sg.pin(
                sg.Column(
                    [[
                        sg.Text("Sheet:", font=FONT_MAIN, text_color=TEXT_DIM,
                                background_color=BG, size=(12, 1)),
                        sg.Combo([], key=f"-{slot}_SHEET-", font=FONT_MAIN,
                                 size=(30, 1), readonly=True,
                                 background_color=BG_CARD, text_color=TEXT),
                        note("  ← select the sheet containing GHG data"),
                    ]],
                    background_color=BG,
                    key=f"-{slot}_SHEET_ROW-",
                    visible=False,
                ),
            ),
        ],
    ]


def _build_layout(config: dict) -> list:
    analyst = config.get("analyst_name", "")
    recent = config.get("recent_projects", [])
    recent_names = [p["name"] for p in recent] if recent else []

    return [
        [sg.Text("GHG Excel QA Tool", font=("Segoe UI", 18, "bold"),
                 text_color=ACCENT, background_color=BG)],
        [note("Compare two versions of a GHG workbook and surface what changed.")],
        [divider()],

        # ── Analyst identity ──────────────────────────────────────────────
        [section_header("Who are you?")],
        [
            label("Your name:", size=(14, 1)),
            sg.Input(analyst, key="-ANALYST-", font=FONT_MAIN, size=(30, 1),
                     background_color=BG_CARD, text_color=TEXT),
        ],
        [sg.Text("", background_color=BG)],  # spacer

        # ── Project ───────────────────────────────────────────────────────
        [section_header("What project is this?")],
        [
            label("Project name:", size=(14, 1)),
            sg.Input(key="-PROJECT-", font=FONT_MAIN, size=(30, 1),
                     background_color=BG_CARD, text_color=TEXT),
            *(
                [
                    sg.Text("  Recent:", font=FONT_SMALL, text_color=TEXT_DIM,
                            background_color=BG),
                    sg.Combo(recent_names, key="-RECENT-", font=FONT_SMALL,
                             size=(20, 1), readonly=True,
                             background_color=BG_CARD, text_color=TEXT,
                             enable_events=True),
                ]
                if recent_names else []
            ),
        ],
        [sg.Text("", background_color=BG)],

        # ── Version 1 ─────────────────────────────────────────────────────
        [section_header("Version 1 (the earlier / baseline version)")],
        *_file_row("V1", "File:"),
        [sg.Text("", background_color=BG)],

        # ── Version 2 ─────────────────────────────────────────────────────
        [section_header("Version 2 (the newer version to compare against)")],
        *_file_row("V2", "File:"),
        [sg.Text("", background_color=BG)],

        # ── Template (optional) ───────────────────────────────────────────
        [section_header("Column Template (optional)")],
        [
            label("Template file:", size=(14, 1)),
            sg.Input(key="-TEMPLATE_PATH-", font=FONT_MAIN, size=(45, 1),
                     background_color=BG_CARD, text_color=TEXT),
            sg.FileBrowse("Browse…", target="-TEMPLATE_PATH-",
                          file_types=(("Excel Files", "*.xlsx"),),
                          font=FONT_SMALL,
                          button_color=(TEXT, BORDER)),
        ],
        [note("If provided, columns will be validated against the template's first sheet headers.")],
        [note("Leave blank to skip — the tool will proceed with best-effort column matching.")],
        [sg.Text("", background_color=BG)],

        [divider()],
        [
            primary_button("Analyze →", "-ANALYZE-"),
            sg.Text("  ", background_color=BG),
            secondary_button("Quit", "-QUIT-"),
        ],
        [sg.Text("", key="-STATUS-", font=FONT_SMALL, text_color=TEXT_DIM,
                 background_color=BG, size=(70, 1))],
    ]


# ---------------------------------------------------------------------------
# Main screen function
# ---------------------------------------------------------------------------

def show_project_screen() -> dict | None:
    """
    Display Screen 1: Project Setup.
    Returns a dict of setup values, or None if the user quits.
    """
    apply_theme()
    config = _load_config()

    layout = _build_layout(config)
    window = sg.Window(
        "GHG QA Tool — Project Setup",
        layout,
        background_color=BG,
        finalize=True,
        resizable=True,
        margins=(20, 20),
    )

    v1_sheets: list[str] = []
    v2_sheets: list[str] = []

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, "-QUIT-"):
            window.close()
            return None

        # ── Recent project selected ───────────────────────────────────────
        if event == "-RECENT-" and values["-RECENT-"]:
            recent = config.get("recent_projects", [])
            match = next((p for p in recent if p["name"] == values["-RECENT-"]), None)
            if match:
                window["-PROJECT-"].update(match["name"])

        # ── File loaded: V1 ───────────────────────────────────────────────
        if event == "-V1_PATH-":
            path = values["-V1_PATH-"].strip()
            if path:
                # Validate extension
                if not path.lower().endswith(".xlsx"):
                    error_popup("Only .xlsx files are supported.\nSave the file as .xlsx in Excel and try again.")
                    window["-V1_PATH-"].update("")
                else:
                    v1_sheets = _load_file_sheets(path, "V1", window)
                    window["-STATUS-"].update(
                        f"V1 loaded: {Path(path).name}  ({len(v1_sheets)} sheet(s))"
                    )

        # ── File loaded: V2 ───────────────────────────────────────────────
        if event == "-V2_PATH-":
            path = values["-V2_PATH-"].strip()
            if path:
                if not path.lower().endswith(".xlsx"):
                    error_popup("Only .xlsx files are supported.\nSave the file as .xlsx in Excel and try again.")
                    window["-V2_PATH-"].update("")
                else:
                    v2_sheets = _load_file_sheets(path, "V2", window)
                    window["-STATUS-"].update(
                        f"V2 loaded: {Path(path).name}  ({len(v2_sheets)} sheet(s))"
                    )

        # ── Analyze clicked ───────────────────────────────────────────────
        if event == "-ANALYZE-":
            analyst = values["-ANALYST-"].strip()
            project = values["-PROJECT-"].strip()
            v1_path = values["-V1_PATH-"].strip()
            v1_label = values["-V1_LABEL-"].strip() or "Version 1"
            v1_sheet = values.get("-V1_SHEET-", "")
            v2_path = values["-V2_PATH-"].strip()
            v2_label = values["-V2_LABEL-"].strip() or "Version 2"
            v2_sheet = values.get("-V2_SHEET-", "")
            template_path = values["-TEMPLATE_PATH-"].strip() or None

            # ── Validation ────────────────────────────────────────────────
            errors = []
            if not analyst:
                errors.append("• Please enter your name.")
            if not project:
                errors.append("• Please enter a project name.")
            if not v1_path:
                errors.append("• Please select the Version 1 file.")
            if not v2_path:
                errors.append("• Please select the Version 2 file.")
            if v1_path and v2_path and Path(v1_path).resolve() == Path(v2_path).resolve():
                errors.append("• Both files appear to be the same. Please select two different versions.")
            if v1_path and not v1_sheet:
                errors.append("• Please select a sheet for Version 1.")
            if v2_path and not v2_sheet:
                errors.append("• Please select a sheet for Version 2.")

            if errors:
                error_popup("\n".join(errors))
                continue

            # ── Save config ───────────────────────────────────────────────
            config["analyst_name"] = analyst
            recent = config.get("recent_projects", [])
            # Upsert project in recent list (keep last 10)
            recent = [p for p in recent if p["name"] != project]
            recent.insert(0, {"name": project, "folder": str(Path(v1_path).parent)})
            config["recent_projects"] = recent[:10]
            _save_config(config)

            window.close()
            return {
                "analyst": analyst,
                "project": project,
                "v1_path": v1_path,
                "v1_label": v1_label,
                "v1_sheet": v1_sheet,
                "v2_path": v2_path,
                "v2_label": v2_label,
                "v2_sheet": v2_sheet,
                "template_path": template_path,
            }
