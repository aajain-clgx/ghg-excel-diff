"""
main.py
-------
Entry point for the GHG Excel QA Tool.
Orchestrates the three-screen flow:

  Screen 1 (Project Setup)
      ↓
  Load sheets + run column matching
      ↓
  Screen 2 (Column Matching Confirmation)
      ↓
  Run diff engine
      ↓
  Screen 3 (Diff Results + Annotation + Export)

The analyst can navigate back to Screen 1 from any later screen.
"""

from __future__ import annotations

import sys
import traceback
from datetime import datetime
from pathlib import Path

# ── Make sure `app/` is on the path when running from project root ────────
sys.path.insert(0, str(Path(__file__).parent))

import FreeSimpleGUI as sg

from diff_engine import load_sheet, run_diff, load_template_headers
from column_matcher import match_columns
from annotation_store import AnnotationStore, sidecar_path
from ui.theme import apply_theme, BG, BG_CARD, TEXT, ACCENT, FONT_MAIN, WHITE, BORDER
from ui.screen_project import show_project_screen
from ui.screen_columns import show_columns_screen
from ui.screen_results import show_results_screen


# ---------------------------------------------------------------------------
# Progress / loading popup
# ---------------------------------------------------------------------------

def _show_loading(message: str) -> sg.Window:
    apply_theme()
    layout = [
        [sg.Text(message, font=FONT_MAIN, text_color=TEXT,
                 background_color=BG_CARD, pad=(20, 20))],
        [sg.ProgressBar(100, orientation="h", size=(30, 10),
                        key="-PROG-", bar_color=(ACCENT, BG_CARD))],
    ]
    w = sg.Window("Loading…", layout, background_color=BG_CARD,
                  no_titlebar=False, keep_on_top=True, finalize=True,
                  margins=(20, 20))
    w["-PROG-"].update_bar(50)
    return w


# ---------------------------------------------------------------------------
# Main orchestration loop
# ---------------------------------------------------------------------------

def main() -> None:
    while True:
        # ── Screen 1: Project Setup ──────────────────────────────────────
        setup = show_project_screen()
        if setup is None:
            break  # User quit

        # ── Load sheets ──────────────────────────────────────────────────
        loading_win = _show_loading(
            f"Loading sheets…\n"
            f"V1: {Path(setup['v1_path']).name}\n"
            f"V2: {Path(setup['v2_path']).name}"
        )

        try:
            df_v1 = load_sheet(setup["v1_path"], setup["v1_sheet"])
            df_v2 = load_sheet(setup["v2_path"], setup["v2_sheet"])

            template_headers = None
            if setup.get("template_path"):
                try:
                    template_headers = load_template_headers(setup["template_path"])
                except Exception:
                    template_headers = None  # Template failure is advisory; don't block

        except PermissionError as exc:
            loading_win.close()
            sg.popup_error(str(exc), title="File Error", font=FONT_MAIN,
                           background_color=BG_CARD, text_color=TEXT)
            continue
        except Exception as exc:
            loading_win.close()
            sg.popup_error(
                f"Could not load files:\n\n{exc}\n\n"
                "Make sure the files are valid .xlsx files and are not open in Excel.",
                title="Load Error", font=FONT_MAIN,
                background_color=BG_CARD, text_color=TEXT,
            )
            continue
        finally:
            loading_win.close()

        # ── Column matching ──────────────────────────────────────────────
        match_result = match_columns(
            headers_v1=list(df_v1.columns),
            headers_v2=list(df_v2.columns),
            template_headers=template_headers,
        )

        # ── Screen 2: Column Matching Confirmation ───────────────────────
        col_setup = show_columns_screen(
            match_result=match_result,
            df_v1=df_v1,
            df_v2=df_v2,
            v1_label=setup["v1_label"],
            v2_label=setup["v2_label"],
        )
        if col_setup is None:
            continue  # Back to Screen 1

        # ── Run diff ─────────────────────────────────────────────────────
        loading_win = _show_loading(
            f"Running diff on {len(df_v1)} sites × "
            f"{len(col_setup['column_mapping'])} columns…"
        )
        try:
            diff_result = run_diff(
                df_v1=df_v1,
                df_v2=df_v2,
                id_column=col_setup["id_column"],
                column_mapping=col_setup["column_mapping"],
                materiality_threshold=col_setup["materiality_threshold"],
            )
        except Exception as exc:
            loading_win.close()
            sg.popup_error(
                f"Diff failed:\n\n{exc}",
                title="Diff Error", font=FONT_MAIN,
                background_color=BG_CARD, text_color=TEXT,
            )
            continue
        finally:
            loading_win.close()

        # ── Set up annotation store ──────────────────────────────────────
        run_date = datetime.now().strftime("%Y-%m-%d")
        project_folder = Path(setup["v1_path"]).parent
        sidecar = sidecar_path(
            project_folder, run_date,
            setup["v1_label"], setup["v2_label"],
        )
        structural = {
            "sites_added":      diff_result.structural.sites_added,
            "sites_removed":    diff_result.structural.sites_removed,
            "columns_added":    diff_result.structural.columns_added,
            "columns_removed":  diff_result.structural.columns_removed,
        }
        store = AnnotationStore.load_or_create(
            sidecar,
            analyst=setup["analyst"],
            project=setup["project"],
            v1_path=setup["v1_path"],
            v1_label=setup["v1_label"],
            v2_path=setup["v2_path"],
            v2_label=setup["v2_label"],
            id_column=col_setup["id_column"],
            materiality_threshold=col_setup["materiality_threshold"],
            column_mapping=col_setup["column_mapping"],
            structural_changes=structural,
        )

        # ── Screen 3: Diff Results ────────────────────────────────────────
        show_results_screen(
            diff_result=diff_result,
            store=store,
            v1_label=setup["v1_label"],
            v2_label=setup["v2_label"],
            project=setup["project"],
        )
        # After Screen 3 closes, loop back to Screen 1 (new comparison)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Last-resort error display — never show a raw traceback to analyst
        apply_theme()
        sg.popup_error(
            "An unexpected error occurred:\n\n" + traceback.format_exc(),
            title="Unexpected Error",
            font=FONT_MAIN,
            background_color=BG_CARD,
            text_color=TEXT,
        )
        sys.exit(1)
