"""
theme.py
--------
Shared colors, fonts, and helper layout elements for the GHG QA Tool UI.
"""

import FreeSimpleGUI as sg

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
BG = "#1E1E2E"           # Main background
BG_CARD = "#2A2A3E"      # Card / panel background
ACCENT = "#4F9CF9"       # Primary action blue
TEXT = "#E0E0F0"         # Primary text
TEXT_DIM = "#8888AA"     # Secondary / label text
RED = "#FF6B6B"          # HIGH severity
YELLOW = "#FFD93D"       # WARN severity
GREEN = "#6BCB77"        # MINOR / OK
WHITE = "#FFFFFF"
BORDER = "#3A3A5C"

FONT_MAIN = ("Segoe UI", 11)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_LARGE = ("Segoe UI", 14, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO = ("Consolas", 10)

# Severity icon + colour lookup
SEVERITY_STYLE = {
    "HIGH":  ("🔴", RED),
    "WARN":  ("🟡", YELLOW),
    "MINOR": ("🟢", GREEN),
}


def apply_theme() -> None:
    """Register and activate the GHG tool colour theme."""
    sg.theme_add_new("GHGDark", {
        "BACKGROUND": BG,
        "TEXT": TEXT,
        "INPUT": BG_CARD,
        "TEXT_INPUT": TEXT,
        "SCROLL": BORDER,
        "BUTTON": (WHITE, ACCENT),
        "PROGRESS": (ACCENT, BORDER),
        "BORDER": 1,
        "SLIDER_DEPTH": 0,
        "PROGRESS_DEPTH": 0,
    })
    sg.theme("GHGDark")


def section_header(title: str) -> sg.Text:
    return sg.Text(title, font=FONT_LARGE, text_color=ACCENT, background_color=BG)


def label(text: str, **kwargs) -> sg.Text:
    return sg.Text(text, font=FONT_MAIN, text_color=TEXT_DIM,
                   background_color=kwargs.pop("background_color", BG), **kwargs)


def note(text: str) -> sg.Text:
    return sg.Text(text, font=FONT_SMALL, text_color=TEXT_DIM, background_color=BG)


def divider(width: int = 80) -> sg.Text:
    return sg.Text("─" * width, text_color=BORDER, background_color=BG, font=FONT_SMALL)


def primary_button(text: str, key: str, **kwargs) -> sg.Button:
    return sg.Button(text, key=key, font=FONT_BOLD,
                     button_color=(WHITE, ACCENT),
                     border_width=0, **kwargs)


def secondary_button(text: str, key: str, **kwargs) -> sg.Button:
    return sg.Button(text, key=key, font=FONT_MAIN,
                     button_color=(TEXT, BORDER),
                     border_width=0, **kwargs)


def error_popup(message: str) -> None:
    sg.popup_error(message, title="Error", font=FONT_MAIN,
                   background_color=BG_CARD, text_color=TEXT,
                   button_color=(WHITE, "#CC3333"))


def info_popup(message: str, title: str = "Info") -> None:
    sg.popup_ok(message, title=title, font=FONT_MAIN,
                background_color=BG_CARD, text_color=TEXT,
                button_color=(WHITE, ACCENT))


def confirm_popup(message: str, title: str = "Confirm") -> bool:
    """Returns True if user clicks Yes."""
    result = sg.popup_yes_no(message, title=title, font=FONT_MAIN,
                             background_color=BG_CARD, text_color=TEXT,
                             button_color=(WHITE, ACCENT))
    return result == "Yes"
