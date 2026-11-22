#!/usr/bin/env python3
"""
Debt Payoff Planner - Desktop App
CustomTkinter application for bi-weekly / semi-monthly / monthly debt payoff planning
with snowball extras, year archiving, and Excel export.
Styled with professional colors matching the original Excel layout.
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import tkinter.font as tkfont
import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
import json
import math
import datetime
from datetime import timedelta
from copy import deepcopy
import os
import sys
import threading
import time
import urllib.request
import webbrowser

# Windows keeps Segoe UI. Linux does not ship that face, so the same
# tuples would fall back to a mismatched font and change the layout.
UI_FONT_FAMILY = "Segoe UI" if sys.platform.startswith("win") else "DejaVu Sans"


def choose_ui_font_family(root=None):
    """Pick a font that is installed. Windows stays on Segoe UI."""
    global UI_FONT_FAMILY
    if sys.platform.startswith("win"):
        UI_FONT_FAMILY = "Segoe UI"
        return UI_FONT_FAMILY
    preferred = ("Noto Sans", "DejaVu Sans", "Liberation Sans", "Ubuntu", "Cantarell", "FreeSans")
    available = set()
    if root is not None:
        try:
            available = set(tkfont.families(root))
        except tk.TclError:
            available = set()
    for name in preferred:
        if name in available:
            UI_FONT_FAMILY = name
            return UI_FONT_FAMILY
    UI_FONT_FAMILY = "DejaVu Sans"
    return UI_FONT_FAMILY


def ui_font(size, *styles):
    if styles:
        return (UI_FONT_FAMILY, size, *styles)
    return (UI_FONT_FAMILY, size)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

DEFAULT_SETTINGS = {
    "frequency": "biweekly",
    "first_pay_date": "2026-10-02",
    "base_extra": 250.0,
    "starting_additional": 0.0,
    "projection_periods": 120,
    "yearly_net_salary": 0.0,
    "spouse_salary": 0.0,
    "other_income": 0.0,
    "spouse_income_frequency": "biweekly",
    "other_income_frequency": "biweekly",
    "income_entered_as_yearly": True,
    "appearance": "dark",
}

APP_NAME = "Debt Payoff Planner"
APP_VERSION = "2.21"
# Startup and File → Check for Updates compare this with version.txt on main.
GITHUB_REPO = "https://github.com/nlaero/DebtPlanner"
UPDATE_VERSION_URL = "https://raw.githubusercontent.com/nlaero/DebtPlanner/main/version.txt"


def parse_version(text):
    """Dot-separated numbers, such as 2.1 or 2.1.1. Anything else is not a version."""
    raw = str(text or "").strip()
    if not raw:
        return None
    token = raw.splitlines()[0].strip()
    if not token or len(token) > 20 or token[0] == "." or token[-1] == "." or ".." in token:
        return None
    parts = []
    for piece in token.split("."):
        if not piece.isdigit():
            return None
        parts.append(int(piece))
    return tuple(parts)


def version_is_newer(remote, local):
    remote_parts = parse_version(remote)
    local_parts = parse_version(local)
    if not remote_parts or not local_parts:
        return False
    width = max(len(remote_parts), len(local_parts))
    remote_parts += (0,) * (width - len(remote_parts))
    local_parts += (0,) * (width - len(local_parts))
    return remote_parts > local_parts


def update_page_url(version):
    parsed = parse_version(version)
    if not parsed:
        return GITHUB_REPO + "#install"
    token = ".".join(str(part) for part in parsed)
    if sys.platform.startswith("win"):
        return (
            "https://github.com/nlaero/DebtPlanner/raw/main/installer/Version%202.x/"
            f"DebtPayoffPlannerSetup-{token}.exe"
        )
    return GITHUB_REPO + "#install"


def fetch_published_version(timeout=8):
    request = urllib.request.Request(
        UPDATE_VERSION_URL,
        headers={"User-Agent": "DebtPayoffPlanner", "Cache-Control": "no-cache"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = response.read(80)
    text = payload.decode("utf-8", errors="replace")
    if parse_version(text) is None:
        raise ValueError("GitHub did not return a version number.")
    return text.strip().splitlines()[0].strip()

# Light is the original palette. Dark follows the slate theme, with a few
# extra shades so amounts and chart lines stay readable.
LIGHT_COLORS = {
    "bg": "#f0f2f5",
    "surface": "#f0f2f5",
    "field": "#ffffff",
    "header_bg": "#1a365d",
    "header_fg": "#ffffff",
    "section_title": "#1a365d",
    "text": "#1a1a1a",
    "text_muted": "#5a6a7a",
    "button": "#1a365d",
    "button_hover": "#234876",
    "button_text": "#ffffff",
    "button_secondary": "#ffffff",
    "button_secondary_hover": "#e4e8ee",
    "button_secondary_text": "#1a365d",
    "button_secondary_border": "#1a365d",
    "green": "#a8e6cf",
    "green_text": "#146c43",
    "paid_selected_bg": "#146c43",
    "paid_selected_text": "#ffffff",
    "extra_text": "#059669",
    "alt_row": "#e8eef4",
    "border": "#bdc3c7",
    "gridline": "#d5dbe3",
    "gridline_on": False,
    "current_row": "#e3f2fd",
    "current_outline": "",
    "total_bg": "#aed6f1",
    "total_fg": "#1a1a1a",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "tab_strip": "#d9dee6",
    "tab_active": "#f0f2f5",
    "tab_active_hover": "#e4e8ee",
    "tab_inactive": "#d9dee6",
    "tab_inactive_hover": "#e7ecf1",
    "tab_active_text": "#1a365d",
    "tab_inactive_text": "#1a365d",
    "tab_edge": "#bdc3c7",
    "tab_line": "#1a365d",
    "schedule_bg": "#ffffff",
    "chart_bg": "#ffffff",
    "chart_plot": "#fbfcfe",
    "chart_grid": "#e6eef5",
    "chart_axis": "#1a365d",
    "chart_label": "#243140",
    "chart_frame": "#d5dee7",
    "chart_wash": "#ffffff",
    "chart_lift": False,
    "legend_off": "#8aa0b2",
    "legend_ring": "#b7c3ce",
    "dropdown_hover": "#e8eef4",
    "snow_bg": "#000000",
    "snow_label": "#d5d5d5",
    "snow_value": "#3DDC67",
}

DARK_COLORS = {
    "bg": "#0F172A",
    "surface": "#1E293B",
    "field": "#1E293B",
    "header_bg": "#1E293B",
    "header_fg": "#F8FAFC",
    "section_title": "#F8FAFC",
    "text": "#F8FAFC",
    "text_muted": "#94A3B8",
    "button": "#3B82F6",
    "button_hover": "#2563EB",
    "button_text": "#F8FAFC",
    "button_secondary": "#334155",
    "button_secondary_hover": "#475569",
    "button_secondary_text": "#F8FAFC",
    "button_secondary_border": "#334155",
    "green": "#10B981",
    "green_text": "#042F2E",
    "paid_selected_bg": "#065F46",
    "paid_selected_text": "#F8FAFC",
    "extra_text": "#34D399",
    "alt_row": "#162033",
    "border": "#334155",
    "gridline": "#1E293B",
    "gridline_on": True,
    "current_row": "#1E293B",
    "current_outline": "#3B82F6",
    "total_bg": "#1E293B",
    "total_fg": "#10B981",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "tab_strip": "#0F172A",
    "tab_active": "#1E293B",
    "tab_active_hover": "#334155",
    "tab_inactive": "#172033",
    "tab_inactive_hover": "#243044",
    "tab_active_text": "#F8FAFC",
    "tab_inactive_text": "#94A3B8",
    "tab_edge": "#334155",
    "tab_line": "#3B82F6",
    "schedule_bg": "#0F172A",
    "chart_bg": "#0F172A",
    "chart_plot": "#1E293B",
    "chart_grid": "#334155",
    "chart_axis": "#94A3B8",
    "chart_label": "#94A3B8",
    "chart_frame": "#334155",
    "chart_wash": "#1E293B",
    "chart_lift": True,
    "legend_off": "#94A3B8",
    "legend_ring": "#475569",
    "dropdown_hover": "#334155",
    "snow_bg": "#1E293B",
    "snow_label": "#94A3B8",
    "snow_value": "#10B981",
}

COLORS = {}


def apply_palette(mode):
    """Copy the light or dark palette into COLORS. Widgets read COLORS when drawn."""
    chosen = "dark" if mode == "dark" else "light"
    COLORS.clear()
    COLORS.update(DARK_COLORS if chosen == "dark" else LIGHT_COLORS)
    return chosen


apply_palette("dark")

_TITLE_ROLES = {
    "Header.TLabel",
    "OverviewValue.TLabel",
    "OverviewHero.TLabel",
}
_MUTED_ROLES = {"OverviewCaption.TLabel"}


def _section_title_chrome():
    """Card titles use the spreadsheet header. In dark mode that header matches the card, so the bar uses the app shell instead."""
    band = COLORS["header_bg"]
    ink = COLORS["header_fg"]
    if band.lower() == COLORS["surface"].lower():
        band = COLORS["bg"]
    rule = COLORS["tab_line"] or ink
    if rule.lower() == band.lower():
        rule = ink
    return band, ink, rule


def _section_note_color():
    """Subtitle on the title band. Light mode uses a pale line on the navy bar."""
    band, _ink, _rule = _section_title_chrome()
    if band.lower() == COLORS["bg"].lower():
        return COLORS["text_muted"]
    return "#d6e4f0"


class AppFrame(ctk.CTkFrame):
    """Page and row frame. padding insets it the way a ttk frame did."""

    def __init__(self, master, padding=0, **kwargs):
        super().__init__(master, fg_color=COLORS["bg"], corner_radius=0, border_width=0)
        self._role = "page"
        self._content_pad = padding

    def pack(self, **kwargs):
        if self._content_pad and "padx" not in kwargs and "pady" not in kwargs:
            kwargs["padx"] = self._content_pad
            kwargs["pady"] = self._content_pad
        return super().pack(**kwargs)

    def grid(self, **kwargs):
        if self._content_pad and "padx" not in kwargs and "pady" not in kwargs:
            kwargs["padx"] = self._content_pad
            kwargs["pady"] = self._content_pad
        return super().grid(**kwargs)


class AppSection(ctk.CTkFrame):
    """Titled card. This object is the body; grid and pack place the outer card."""

    def __init__(self, master, text="", padding=10, style=None, note="", **kwargs):
        self.shell = ctk.CTkFrame(
            master,
            fg_color=COLORS["surface"],
            corner_radius=10,
            border_width=1,
            border_color=COLORS["border"],
        )
        self.shell._role = "card"
        band_color, ink, rule_color = _section_title_chrome()
        band = ctk.CTkFrame(self.shell, fg_color=band_color, corner_radius=8)
        band._role = "section_band"
        band.pack(fill="x", padx=8, pady=(8, 0))
        title = ctk.CTkLabel(
            band,
            text=text,
            font=ui_font( 18, "bold"),
            text_color=ink,
            fg_color="transparent",
            anchor="w",
            height=28,
        )
        title._role = "section"
        title.pack(fill="x", padx=12, pady=(6, 0))
        self._title_note = None
        if note:
            note_label = ctk.CTkLabel(
                band,
                text=note,
                font=ui_font(12),
                text_color=_section_note_color(),
                fg_color="transparent",
                anchor="w",
                justify="left",
                height=18,
            )
            note_label._role = "section_note"
            note_label.pack(fill="x", padx=12, pady=(0, 0))
            self._title_note = note_label
        rule = ctk.CTkFrame(band, fg_color=rule_color, height=3, corner_radius=2)
        rule._role = "section_rule"
        rule.pack(fill="x", padx=12, pady=(4, 8))
        self._title_band = band
        self._title = title
        self._title_rule = rule
        super().__init__(self.shell, fg_color="transparent", corner_radius=0, border_width=0)
        self._role = "clear"
        ctk.CTkFrame.pack(self, fill="both", expand=True, padx=padding, pady=(8, padding))

    def grid(self, **kwargs):
        return self.shell.grid(**kwargs)

    def pack(self, **kwargs):
        return self.shell.pack(**kwargs)

    def grid_remove(self):
        return self.shell.grid_remove()

    def pack_forget(self):
        return self.shell.pack_forget()


class AppLabel(ctk.CTkLabel):
    def __init__(self, master, text=None, textvariable=None, style=None, anchor="w",
                 justify="left", wraplength=0, background=None, foreground=None, relief=None,
                 role=None, **kwargs):
        self._role = role or style or "body"
        fonts = {
            "Header.TLabel": ui_font( 13, "bold"),
            "OverviewValue.TLabel": ui_font( 13, "bold"),
            "OverviewCaption.TLabel": ui_font( 12),
            "OverviewHero.TLabel": ui_font( 19, "bold"),
        }
        colors = {
            "Header.TLabel": COLORS["section_title"],
            "OverviewValue.TLabel": COLORS["section_title"],
            "OverviewCaption.TLabel": COLORS["text_muted"],
            "OverviewHero.TLabel": COLORS["section_title"],
            "status": COLORS["header_fg"],
        }
        heights = {
            "OverviewHero.TLabel": 34,
            "OverviewCaption.TLabel": 24,
        }
        content = "" if text is None else text
        height = heights.get(style, 28)
        if "\n" in content:
            height = max(height, content.count("\n") * 18 + 28)
            if not wraplength:
                wraplength = 980
        super().__init__(
            master,
            text=content,
            textvariable=textvariable,
            font=fonts.get(style, ui_font( 12)),
            text_color=colors.get(style, foreground or COLORS["text"]),
            fg_color=background or "transparent",
            anchor=anchor or "w",
            justify=justify,
            wraplength=wraplength,
            height=height,
            corner_radius=0,
        )


class AppButton(ctk.CTkButton):
    def __init__(self, master, text="", command=None, style=None, **kwargs):
        compact = style == "ChartKey.TButton"
        secondary = style in ("secondary", "ChartKey.TButton")
        if style == "danger":
            self._button_tone = "danger"
        elif secondary:
            self._button_tone = "secondary"
        else:
            self._button_tone = "primary"
        label = "" if text is None else str(text)
        width = max(96, min(420, len(label) * 8 + 36))
        if style == "danger":
            fill = COLORS["danger"]
            hover = COLORS["danger_hover"]
            text_color = COLORS["button_text"]
            border = fill
        elif secondary:
            fill = COLORS["button_secondary"]
            hover = COLORS["button_secondary_hover"]
            text_color = COLORS["button_secondary_text"]
            border = COLORS["button_secondary_border"]
        else:
            fill = COLORS["button"]
            hover = COLORS["button_hover"]
            text_color = COLORS["button_text"]
            border = fill
        super().__init__(
            master,
            text=label,
            command=command,
            fg_color=fill,
            hover_color=hover,
            text_color=text_color,
            border_width=1 if secondary else 0,
            border_color=border,
            corner_radius=6,
            height=28 if compact else 32,
            width=width if not compact else max(96, width),
            font=ui_font( 12 if compact else 13, "bold"),
        )


class AppEntry(ctk.CTkEntry):
    def __init__(self, master, textvariable=None, width=20, **kwargs):
        super().__init__(
            master,
            textvariable=textvariable,
            width=max(90, int(width) * 8),
            height=32,
            font=ui_font( 13),
            fg_color=COLORS["field"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            corner_radius=6,
        )


class AppCombo(ctk.CTkComboBox):
    def __init__(self, master, textvariable=None, values=None, state="readonly", width=20, **kwargs):
        super().__init__(
            master,
            variable=textvariable,
            values=list(values) if values else [],
            state="readonly" if state == "readonly" else state,
            width=max(150, int(width) * 8 + 40),
            height=32,
            font=ui_font( 13),
            fg_color=COLORS["field"],
            border_color=COLORS["border"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            dropdown_fg_color=COLORS["field"],
            dropdown_hover_color=COLORS["dropdown_hover"],
            dropdown_text_color=COLORS["text"],
            text_color=COLORS["text"],
            corner_radius=6,
        )
        self._select_handlers = []

    def __setitem__(self, key, value):
        if key == "values":
            self.configure(values=list(value))
        else:
            self.configure(**{key: value})

    def __getitem__(self, key):
        return self.cget(key)

    def bind(self, sequence=None, command=None, add=True):
        if sequence == "<<ComboboxSelected>>":
            self._select_handlers.append(command)
            self.configure(command=self._fire_select)
            return "combo-select"
        return super().bind(sequence, command, add=True if add is None else add)

    def _fire_select(self, _value):
        event = tk.Event()
        event.widget = self
        for handler in self._select_handlers:
            handler(event)


class AppCheck(ctk.CTkCheckBox):
    def __init__(self, master, text="", variable=None, **kwargs):
        super().__init__(
            master,
            text=text,
            variable=variable,
            onvalue=True,
            offvalue=False,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            checkmark_color=COLORS["button_text"],
            font=ui_font( 12),
        )


class AppScroll(ctk.CTkScrollbar):
    def __init__(self, master, orient="vertical", command=None, **kwargs):
        super().__init__(
            master,
            orientation="horizontal" if str(orient).startswith("h") else "vertical",
            command=command,
            button_color=COLORS["border"],
            button_hover_color=COLORS["header_bg"],
            fg_color=COLORS["bg"],
        )


# ---------------------------------------------------------------------------
# Core calculation helpers
# ---------------------------------------------------------------------------

def _key_variants(key):
    variants = [key]
    if not isinstance(key, str):
        variants.append(str(key))
    try:
        as_int = int(key)
    except (TypeError, ValueError):
        as_int = None
    if as_int is not None and as_int not in variants:
        variants.append(as_int)
    return variants


def mapping_get(mapping, key, default=None):
    """Look up a key that may be an int in memory and a string after JSON."""
    if not isinstance(mapping, dict):
        return default
    for variant in _key_variants(key):
        if variant in mapping:
            return mapping[variant]
    return default


def normalize_id_map(raw):
    """Restore int keys after JSON turns object keys into strings."""
    if not isinstance(raw, dict):
        return {}
    out = {}
    for key, value in raw.items():
        try:
            normalized_key = int(key)
        except (TypeError, ValueError):
            normalized_key = key
        if isinstance(value, dict):
            inner = {}
            for child_key, child_value in value.items():
                try:
                    inner[int(child_key)] = child_value
                except (TypeError, ValueError):
                    inner[child_key] = child_value
            out[normalized_key] = inner
        else:
            out[normalized_key] = value
    return out


def _normalize_archive_columns(columns):
    frozen = []
    for column in columns:
        if not isinstance(column, dict) or "id" not in column:
            continue
        try:
            did = int(column["id"])
        except (TypeError, ValueError):
            did = column["id"]
        frozen.append({"id": did, "short_name": column.get("short_name") or f"Debt {did}"})
    return frozen


def normalize_archives(raw):
    if not isinstance(raw, dict):
        return {}
    out = {}
    for year, periods in raw.items():
        try:
            year_key = int(year)
        except (TypeError, ValueError):
            continue
        cleaned = []
        for period in periods or []:
            if not isinstance(period, dict):
                continue
            item = dict(period)
            if isinstance(item.get("balances"), dict):
                item["balances"] = normalize_id_map(item["balances"])
            if isinstance(item.get("extras"), dict):
                item["extras"] = normalize_id_map(item["extras"])
            if isinstance(item.get("closings"), dict):
                item["closings"] = normalize_id_map(item["closings"])
            if isinstance(item.get("columns"), list):
                item["columns"] = _normalize_archive_columns(item["columns"])
            date_key = str(item.get("date") or "")
            if date_key and any(str(existing.get("date") or "") == date_key for existing in cleaned):
                continue
            cleaned.append(item)
        out[year_key] = cleaned
    return out


PAY_FREQUENCIES = ("biweekly", "semimonthly", "monthly")


def periods_per_year(frequency):
    return {"biweekly": 26, "semimonthly": 24, "monthly": 12}.get(frequency, 26)


def income_frequency_key(value):
    text = str(value or "").strip().lower().replace(" ", "").replace("-", "")
    aliases = {
        "biweekly": "biweekly",
        "semimonthly": "semimonthly",
        "monthly": "monthly",
        "yearly": "yearly",
        "annual": "yearly",
        "perpay": "per_pay",
        "perpaycheck": "per_pay",
    }
    return aliases.get(text, "monthly")


def period_amount_for_income(amount, income_frequency):
    """Yearly net split across the selected Overview pay-period frequency."""
    number = parse_amount(amount, minimum=0, allow_blank=True) or 0.0
    how = income_frequency_key(income_frequency)
    if how == "yearly":
        return number / 12.0
    ppy = periods_per_year(how) or 12
    return number / ppy


def _income_already_yearly(raw):
    flag = None
    if isinstance(raw, dict):
        flag = raw.get("income_entered_as_yearly")
    return flag in (True, "true", "True", 1, "1")


def _normalize_income_line(settings, amount_key, freq_key, already_yearly):
    """Keep the Settings amount as yearly net. Dropdown is Overview pay period only."""
    how = income_frequency_key(settings.get(freq_key))
    amount = parse_amount(settings.get(amount_key), minimum=0, allow_blank=True) or 0.0
    pay = settings.get("frequency") if settings.get("frequency") in PAY_FREQUENCIES else "biweekly"
    if already_yearly:
        if how == "yearly":
            settings[freq_key] = "monthly"
        elif how == "per_pay":
            settings[freq_key] = pay
        elif how in PAY_FREQUENCIES:
            settings[freq_key] = how
        else:
            settings[freq_key] = "biweekly"
        settings[amount_key] = amount
        return
    if how == "yearly":
        settings[amount_key] = amount
        settings[freq_key] = "monthly"
    elif how == "per_pay":
        settings[amount_key] = amount * (periods_per_year(pay) or 26)
        settings[freq_key] = pay
    elif how in PAY_FREQUENCIES:
        settings[amount_key] = amount * (periods_per_year(how) or 12)
        settings[freq_key] = how
    else:
        settings[amount_key] = amount * 12.0
        settings[freq_key] = "monthly"

def add_period(date, frequency):
    if frequency == "biweekly":
        return date + timedelta(days=14)
    elif frequency == "semimonthly":
        return date + timedelta(days=15)
    else:
        month = date.month + 1
        year = date.year
        if month > 12:
            month = 1
            year += 1
        day = min(date.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month-1])
        return datetime.date(year, month, day)

def parse_flexible_date(text):
    """Read a stored YYYY-MM-DD date or a typed MM-DD-YYYY date."""
    cleaned = str(text or "").strip().replace("/", "-")
    if not cleaned:
        return None
    if len(cleaned) >= 10 and cleaned[4] == "-" and cleaned[:4].isdigit():
        try:
            return datetime.datetime.strptime(cleaned[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    try:
        return datetime.datetime.strptime(cleaned, "%m-%d-%Y").date()
    except ValueError:
        return None


def format_display_date(value):
    """Show a full date as MM-DD-YYYY. Month-day entries and labels stay as they are."""
    if value is None:
        return ""
    text = str(value).strip()
    if text in ("", "✓", "Never", "—", "-"):
        return text
    parsed = parse_flexible_date(text)
    if parsed is None:
        return text
    return parsed.strftime("%m-%d-%Y")


def parse_date(value):
    return parse_flexible_date(value) or datetime.date.today()


def parse_iso_date(text):
    return parse_flexible_date(text)


def parse_amount(value, minimum=None, maximum=None, allow_blank=False):
    """Finite float, or None if invalid. Blank is 0 when allow_blank is True."""
    if value is None:
        return 0.0 if allow_blank else None
    text = str(value).strip().replace(",", "").replace("$", "")
    if text.endswith("%"):
        text = text[:-1].strip()
    if text in ("", "+", "-", "."):
        return 0.0 if allow_blank else None
    try:
        number = float(text)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number):
        return None
    if minimum is not None and number < minimum:
        return None
    if maximum is not None and number > maximum:
        return None
    return number


def format_money_entry(value):
    """$91,000 for whole dollars, $91,000.12 when there are cents."""
    number = parse_amount(value, allow_blank=True)
    if number is None:
        return None
    sign = "-" if number < 0 else ""
    whole, frac = f"{abs(number):.2f}".split(".")
    if frac == "00":
        return f"{sign}${int(whole):,}"
    return f"{sign}${int(whole):,}.{frac}"


def parse_period_count(value):
    number = parse_amount(value, minimum=1, maximum=2600)
    if number is None:
        return None
    return int(number)


def sanitize_debt(debt):
    if not isinstance(debt, dict) or "id" not in debt:
        return None
    try:
        did = int(debt["id"])
    except (TypeError, ValueError):
        return None
    current = parse_amount(debt.get("current"), minimum=0, maximum=1e12, allow_blank=True)
    original = parse_amount(debt.get("original"), minimum=0, maximum=1e12, allow_blank=True)
    minimum = parse_amount(debt.get("min_payment"), minimum=0, maximum=1e12, allow_blank=True)
    rate = parse_amount(debt.get("rate"), minimum=0, maximum=10, allow_blank=True)
    if None in (current, original, minimum, rate):
        current = current if current is not None else 0.0
        original = original if original is not None else 0.0
        minimum = minimum if minimum is not None else 0.0
        rate = rate if rate is not None else 0.0
    cleaned = dict(debt)
    cleaned["id"] = did
    cleaned["short_name"] = str(debt.get("short_name") or f"Debt {did}").strip() or f"Debt {did}"
    cleaned["full_name"] = str(debt.get("full_name") or cleaned["short_name"]).strip()
    cleaned["current"] = current
    cleaned["original"] = original
    cleaned["min_payment"] = minimum
    cleaned["rate"] = rate
    cleaned["include_paid_minimum"] = include_paid_minimum(debt)
    return cleaned


def sanitize_settings(raw):
    settings = deepcopy(DEFAULT_SETTINGS)
    if not isinstance(raw, dict):
        return settings
    frequency = raw.get("frequency", settings["frequency"])
    if frequency not in ("biweekly", "semimonthly", "monthly"):
        frequency = "biweekly"
    settings["frequency"] = frequency
    first = parse_iso_date(raw.get("first_pay_date"))
    if first:
        settings["first_pay_date"] = first.isoformat()
    balance = parse_iso_date(raw.get("balance_as_of") or raw.get("first_pay_date") or settings["first_pay_date"])
    if balance:
        settings["balance_as_of"] = balance.isoformat()
    else:
        settings["balance_as_of"] = settings["first_pay_date"]
    periods = parse_period_count(raw.get("projection_periods", settings["projection_periods"]))
    settings["projection_periods"] = periods if periods is not None else 120
    for key in ("starting_additional", "yearly_net_salary", "spouse_salary", "other_income", "base_extra"):
        amount = parse_amount(raw.get(key), minimum=0, maximum=1e12, allow_blank=True)
        settings[key] = amount if amount is not None else 0.0
    earliest = parse_iso_date(raw.get("earliest_pay_date"))
    if earliest:
        settings["earliest_pay_date"] = earliest.isoformat()
    settings["spouse_income_frequency"] = raw.get("spouse_income_frequency")
    settings["other_income_frequency"] = raw.get("other_income_frequency")
    already_yearly = _income_already_yearly(raw)
    _normalize_income_line(settings, "spouse_salary", "spouse_income_frequency", already_yearly)
    _normalize_income_line(settings, "other_income", "other_income_frequency", already_yearly)
    settings["income_entered_as_yearly"] = True
    settings["appearance"] = "light" if raw.get("appearance") == "light" else "dark"
    return settings


def archived_pay_dates(archives):
    dates = set()
    if not isinstance(archives, dict):
        return dates
    for periods in archives.values():
        for period in periods or []:
            if not isinstance(period, dict):
                continue
            parsed = parse_iso_date(period.get("date"))
            if parsed:
                dates.add(parsed)
    return dates


def last_archived_pay_date(archives):
    dates = archived_pay_dates(archives)
    return max(dates) if dates else None


def first_friday_of_year(year):
    """First Friday of a calendar year. January 1 is weekday 0 through 6, Monday through Sunday."""
    new_year = datetime.date(int(year), 1, 1)
    return new_year + timedelta(days=(4 - new_year.weekday()) % 7)


def _shift_iso_date(value, delta):
    parsed = parse_iso_date(value)
    if parsed is None:
        return value
    return (parsed + delta).isoformat()


def rebase_sample_calendar(settings, debts, extras_by_period, archives, today=None):
    """Put a packaged sample on the first Friday of the year it is opened.

    The file's payday, archive dates, extras, and due dates stay the same
    distance apart. A new install in a later year still has a current payday,
    a live schedule, and history that sits before that start.
    """
    today = today or datetime.date.today()
    source = parse_iso_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    if source is None:
        return
    target = first_friday_of_year(today.year)
    delta = target - source
    if delta.days == 0:
        return
    for key in ("first_pay_date", "balance_as_of", "earliest_pay_date"):
        if settings.get(key):
            settings[key] = _shift_iso_date(settings[key], delta)
    for debt in debts or []:
        if not isinstance(debt, dict):
            continue
        for key in ("due_anchor", "due_date", "schedule_start"):
            if debt.get(key):
                debt[key] = _shift_iso_date(debt[key], delta)
        anchor = parse_iso_date(debt.get("due_anchor"))
        if anchor is not None:
            debt["due_md"] = f"{anchor.month:02d}-{anchor.day:02d}"
            debt["due_date"] = anchor.isoformat()
    if isinstance(extras_by_period, dict):
        shifted_extras = {}
        for key, value in extras_by_period.items():
            parsed = parse_iso_date(key)
            shifted_extras[_shift_iso_date(key, delta) if parsed else key] = value
        extras_by_period.clear()
        extras_by_period.update(shifted_extras)
    if isinstance(archives, dict):
        rebuilt = {}
        for rows in list(archives.values()):
            if not isinstance(rows, list):
                continue
            for row in rows:
                if not isinstance(row, dict):
                    continue
                parsed = parse_iso_date(row.get("date"))
                if parsed is None:
                    continue
                moved = parsed + delta
                row["date"] = moved.isoformat()
                row["year"] = moved.year
                rebuilt.setdefault(moved.year, []).append(row)
        archives.clear()
        archives.update(rebuilt)


def sanitize_extras(raw):
    if not isinstance(raw, dict):
        return {}
    cleaned = {}
    for key, value in raw.items():
        if not isinstance(value, dict):
            continue
        amounts = {}
        for did, amount in value.items():
            try:
                debt_id = int(did)
            except (TypeError, ValueError):
                continue
            number = parse_amount(amount, minimum=0, maximum=1e12)
            if number and number > 0:
                amounts[debt_id] = number
        if amounts:
            cleaned[str(key)] = amounts
    return cleaned


def subtract_period(date, frequency):
    if frequency == "biweekly":
        return date - timedelta(days=14)
    elif frequency == "semimonthly":
        return date - timedelta(days=15)
    month = date.month - 1
    year = date.year
    if month < 1:
        month = 12
        year -= 1
    day = min(date.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return datetime.date(year, month, day)


def first_aligned_pay_of_year(anchor, frequency):
    """First pay date of anchor's year that lands on the same cadence."""
    current = anchor
    for _ in range(500):
        prev = subtract_period(current, frequency)
        if prev.year != current.year:
            return current
        current = prev
    return current


def align_on_or_after(requested, anchor, frequency):
    """Nearest cadence date on or after requested, counting from anchor."""
    if requested <= anchor:
        current = anchor
        for _ in range(500):
            prev = subtract_period(current, frequency)
            if prev < requested:
                return current
            current = prev
        return current
    current = anchor
    for _ in range(500):
        if current >= requested:
            return current
        current = add_period(current, frequency)
    return current


def period_in_year(date, frequency):
    first = first_aligned_pay_of_year(date, frequency)
    number = 1
    current = first
    for _ in range(500):
        if current >= date:
            return number
        current = add_period(current, frequency)
        number += 1
    return number


def align_on_or_before(requested, anchor, frequency):
    """Nearest cadence date on or before requested, counting from anchor."""
    if requested >= anchor:
        current = anchor
        for _ in range(500):
            nxt = add_period(current, frequency)
            if nxt > requested:
                return current
            current = nxt
        return current
    current = anchor
    for _ in range(500):
        prev = subtract_period(current, frequency)
        if prev <= requested:
            return prev
        current = prev
    return current


def current_pay_date(settings, today=None):
    """Pay date still in effect, using this plan's pay date and frequency.

    A period stays current from its pay date through the day before the next
    one. It changes only when today is that next pay date.
    """
    today = today or datetime.date.today()
    freq = settings.get("frequency", "biweekly")
    anchor = parse_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    return align_on_or_before(today, anchor, freq)


def schedule_start_date(settings, archives=None):
    """Earliest live pay date. Archived paydays are never rebuilt on the schedule."""
    freq = settings.get("frequency", "biweekly")
    anchor = parse_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    requested = parse_date(settings.get("first_pay_date"))
    candidate = first_aligned_pay_of_year(anchor, freq)
    if requested < candidate:
        candidate = align_on_or_after(requested, anchor, freq)
    floors = []
    earliest = parse_iso_date(settings.get("earliest_pay_date"))
    if earliest:
        floors.append(align_on_or_after(earliest, anchor, freq))
    last_archived = last_archived_pay_date(archives)
    if last_archived:
        floors.append(align_on_or_after(last_archived + timedelta(days=1), anchor, freq))
    for floor_date in floors:
        if floor_date > candidate:
            candidate = floor_date
    return candidate


DEBT_TYPES = ("Auto", "Personal", "Mortgage", "Credit", "Payment Plan")
OVERVIEW_TYPES = DEBT_TYPES + ("Unspecified",)
DUE_FREQUENCIES = ("weekly", "biweekly", "semimonthly", "monthly")
DUE_FREQUENCY_LABELS = {
    "weekly": "Weekly",
    "biweekly": "Bi-weekly",
    "semimonthly": "Semi-monthly",
    "monthly": "Monthly",
}
DUE_INTERVAL_DAYS = {"weekly": 7, "biweekly": 14, "semimonthly": 15}
ANNUAL_DUE_COUNTS = {"weekly": 52, "biweekly": 26, "semimonthly": 24, "monthly": 12}


def include_paid_minimum(debt):
    """Roll this debt's regular payment into Additional Per Pay after it is paid off.

    Missing on an older plan means included, which is the default.
    """
    if not isinstance(debt, dict) or "include_paid_minimum" not in debt:
        return True
    flag = debt.get("include_paid_minimum")
    if isinstance(flag, str):
        return flag.strip().lower() in ("1", "true", "yes", "on")
    return bool(flag)


def additional_per_pay(debts, settings):
    """Freed minimums from paid-off debts, spread across one year of paychecks.

    A debt left unchecked does not add its regular payment after payoff.
    """
    annual = 0.0
    for debt in debts:
        balance = parse_amount(debt.get("current"), allow_blank=True)
        minimum = parse_amount(debt.get("min_payment"), allow_blank=True)
        if balance is None or minimum is None:
            continue
        if balance >= 0.5 or minimum <= 0 or not include_paid_minimum(debt):
            continue
        frequency = debt.get("due_frequency") or "monthly"
        annual += minimum * ANNUAL_DUE_COUNTS.get(frequency, 12)
    pay_periods = periods_per_year(settings.get("frequency", "biweekly"))
    if pay_periods <= 0:
        return 0.0
    starting = parse_amount(settings.get("starting_additional"), minimum=0, allow_blank=True) or 0.0
    return round(starting + (annual / pay_periods), 2)


def annual_minimum(debt):
    minimum = parse_amount(debt.get("min_payment"), minimum=0, allow_blank=True)
    if minimum is None:
        return 0.0
    frequency = debt.get("due_frequency") or "monthly"
    return minimum * ANNUAL_DUE_COUNTS.get(frequency, 12)


def _is_active_debt(debt):
    value = parse_amount(debt.get("current"), allow_blank=True)
    return value is not None and value >= 0.5


def _debt_type_name(debt):
    kind = debt.get("debt_type")
    return kind if kind in DEBT_TYPES else "Unspecified"


def _payoff_rank_key(debt):
    try:
        rank = int(debt.get("payoff_rank"))
    except (TypeError, ValueError):
        rank = 10 ** 9
    if rank < 1:
        rank = 10 ** 9
    try:
        did = int(debt.get("id") or 0)
    except (TypeError, ValueError):
        did = 0
    return (rank, did)


def retire_paid_off_debts(debts):
    """Paid-off debts leave ranking and the live schedule; remaining stay in order."""
    for debt in debts:
        if _is_active_debt(debt):
            continue
        debt.pop("payoff_rank", None)
        debt["schedule_hidden"] = True


def assign_payoff_ranks(debts, ordered=None):
    """Give unpaid debts unique ranks 1..n. Paid-off debts are unranked at the bottom."""
    retire_paid_off_debts(debts)
    if ordered is None:
        ordered = [debt for debt in debts if _is_active_debt(debt)]
        ordered.sort(key=_payoff_rank_key)
    else:
        ordered = [debt for debt in ordered if _is_active_debt(debt)]
    seen = set()
    compact = []
    for debt in ordered:
        marker = id(debt)
        if marker in seen:
            continue
        seen.add(marker)
        compact.append(debt)
    for debt in debts:
        if _is_active_debt(debt) and id(debt) not in seen:
            compact.append(debt)
            seen.add(id(debt))
    for index, debt in enumerate(compact, 1):
        debt["payoff_rank"] = index
    return compact


def _period_minimum(debt, minimum, per_pay_min, period_start, next_period):
    """Minimum applied this payday. No Next Due still uses the per-paycheck share."""
    if due_series(debt)[0] is None:
        return per_pay_min
    return minimum * minimum_payments_due(debt, period_start, next_period)


def _category_dates_from_debts(ranked, by_id):
    """A type is debt-free on the latest individual payoff in that type."""
    dates = {name: "✓" for name in OVERVIEW_TYPES}
    grouped = {}
    for debt in ranked:
        grouped.setdefault(_debt_type_name(debt), []).append(by_id.get(debt["id"], "Never"))
    for name, items in grouped.items():
        dates[name] = "Never" if "Never" in items else max(items)
    if not ranked:
        overall = "✓"
    elif "Never" in by_id.values():
        overall = "Never"
    else:
        overall = max(by_id.values()) if by_id else "✓"
    return dates, overall


def _logged_extra_amounts(extras_by_period, pay_text, debt_ids):
    """Extra actually entered on this payday, keyed by debt id."""
    logged = mapping_get(extras_by_period, pay_text, {})
    amounts = {}
    if not isinstance(logged, dict):
        return amounts
    for did in debt_ids:
        amount = parse_amount(mapping_get(logged, did, 0.0), minimum=0, allow_blank=True) or 0.0
        if amount > 0:
            amounts[did] = amount
    return amounts


def debt_free_dates_by_type(debts, settings, today=None, extras_by_period=None):
    """Snowball dates for each debt, then rolled up by debt type.

    Balances start at Balances As Of. Minimums are paid every payday.
    On a payday that has arrived, extra is whatever was logged on the Pay
    Schedule for that date. Nothing logged means that expected extra was
    not paid. A larger logged payment pays the debt off sooner. Later
    paydays that have no logged extra still use Starting Additional Per Pay
    plus freed minimums, in payoff rank order.
    """
    by_id = {}
    ranked = [debt for debt in debts if _is_active_debt(debt)]
    ranked.sort(key=_payoff_rank_key)
    if not ranked:
        dates, overall = _category_dates_from_debts(ranked, by_id)
        return dates, overall, by_id

    today = today or datetime.date.today()
    extras_by_period = extras_by_period or {}
    freq = settings.get("frequency", "biweekly")
    ppy = periods_per_year(freq) or 12
    extra = additional_per_pay(debts, settings)
    start = parse_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    display_periods = parse_period_count(settings.get("projection_periods")) or 120
    max_periods = max(display_periods, ppy * 50)
    rows = []
    balances = {}
    starts = {}
    debt_ids = []
    for debt in ranked:
        did = debt["id"]
        debt_ids.append(did)
        balances[did] = parse_amount(debt.get("current"), minimum=0, allow_blank=True) or 0.0
        started = parse_iso_date(debt.get("schedule_start"))
        if started:
            starts[did] = started
        rows.append((
            did,
            debt,
            parse_amount(debt.get("rate"), minimum=0, allow_blank=True) or 0.0,
            parse_amount(debt.get("min_payment"), minimum=0, allow_blank=True) or 0.0,
            annual_minimum(debt) / ppy,
        ))
    current = start
    for _ in range(max_periods):
        nxt = add_period(current, freq)
        pay_text = current.isoformat()
        logged_amounts = _logged_extra_amounts(extras_by_period, pay_text, debt_ids)
        use_logged = current <= today or bool(logged_amounts)
        after = {}
        for did, debt, rate, minimum, per_pay_min in rows:
            if did in by_id:
                after[did] = 0.0
                continue
            bal = balances.get(did, 0.0)
            if starts.get(did) and current < starts[did]:
                after[did] = bal
                continue
            interest = bal * (rate / ppy) if rate > 0 else 0.0
            after[did] = bal + interest - _period_minimum(debt, minimum, per_pay_min, current, nxt)
        if use_logged:
            for did, debt, rate, minimum, per_pay_min in rows:
                if did in by_id or (starts.get(did) and current < starts[did]):
                    continue
                bal = after.get(did, 0.0)
                if bal < 0.5:
                    after[did] = 0.0
                    by_id[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                    continue
                pay = min(logged_amounts.get(did, 0.0), bal)
                bal -= pay
                if bal < 0.5:
                    after[did] = 0.0
                    by_id[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                else:
                    after[did] = bal
        else:
            pool = extra
            for did, debt, rate, minimum, per_pay_min in rows:
                if did in by_id or (starts.get(did) and current < starts[did]):
                    continue
                bal = after.get(did, 0.0)
                if bal < 0.5:
                    after[did] = 0.0
                    by_id[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                    continue
                if pool <= 0:
                    after[did] = bal
                    continue
                pay = min(pool, bal)
                bal -= pay
                pool -= pay
                if bal < 0.5:
                    after[did] = 0.0
                    by_id[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                else:
                    after[did] = bal
        balances = after
        if len(by_id) >= len(rows):
            break
        current = nxt
    for debt in ranked:
        by_id.setdefault(debt["id"], "Never")
    dates, overall = _category_dates_from_debts(ranked, by_id)
    return dates, overall, by_id


MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
# Spread around the navy, green, and blue already used in the app.
CHART_COLORS = (
    "#1a365d",
    "#922b21",
    "#1e8449",
    "#6c3483",
    "#b9770e",
    "#2471a3",
    "#a84b6a",
    "#4e342e",
)


def chart_mix_toward(color, target, weight):
    """Blend a hex color toward another. weight 0 keeps color, 1 is the target."""
    def channels(hex_color):
        value = str(hex_color).lstrip("#")
        if len(value) != 6:
            return None
        return [int(value[index:index + 2], 16) for index in (0, 2, 4)]

    src = channels(color)
    dst = channels(target)
    if not src or not dst:
        return color
    mixed = [round(channel + (dest - channel) * weight) for channel, dest in zip(src, dst)]
    return "#{:02x}{:02x}{:02x}".format(*mixed)


def chart_display_color(color):
    """Lift dark series colors on the dark chart so the lines stay visible."""
    if COLORS.get("chart_lift"):
        return chart_mix_toward(color, "#F8FAFC", 0.45)
    return color


def chart_tick_label(day, detailed):
    """Month and 'YY, or a pay date when the span is under a year."""
    if detailed:
        return day.strftime("%m/%d/%y")
    return f"{MONTH_NAMES[day.month - 1]} '{day.year % 100:02d}"


def chart_axis_ticks(start, end, frequency, plot_width, detailed=None, min_px=80):
    """Tick positions and labels spaced so neighbors do not run together."""
    if end < start:
        end = start
    if detailed is None:
        detailed = (end - start).days < 365
    if detailed:
        items = []
        current = start
        for _ in range(2000):
            if current > end:
                break
            items.append((current, chart_tick_label(current, True)))
            nxt = add_period(current, frequency)
            if nxt <= current:
                break
            current = nxt
        if not items:
            items = [(start, chart_tick_label(start, True))]
        if items[-1][0] != end:
            items.append((end, chart_tick_label(end, True)))
    else:
        items = []
        year, month = start.year, start.month
        for _ in range(600):
            month_start = datetime.date(year, month, 1)
            if month_start > end and (year != end.year or month != end.month):
                break
            position = start if month_start < start else month_start
            if position > end:
                break
            items.append((position, chart_tick_label(month_start, False)))
            month += 1
            if month > 12:
                month = 1
                year += 1
        if not items:
            items = [(start, chart_tick_label(start, False))]
    if len(items) <= 2:
        return detailed, items
    min_px = max(int(min_px), 24)
    span = max((end - start).days, 1)
    width = max(int(plot_width), 1)

    def at(day):
        return (day - start).days / span * width

    stride = 1
    chosen = items
    while stride <= len(items):
        chosen = list(items[::stride])
        last = items[-1]
        if chosen[-1][0] != last[0] and at(last[0]) - at(chosen[-1][0]) >= min_px:
            chosen.append(last)
        if len(chosen) <= 2:
            break
        gaps = [at(chosen[i][0]) - at(chosen[i - 1][0]) for i in range(1, len(chosen))]
        if gaps and min(gaps) >= min_px * 0.92:
            break
        stride += 1
    return detailed, chosen


def chart_value_ticks(y_max, plot_height, min_px):
    """Dollar ticks spaced so the labels do not sit on each other."""
    y_max = max(float(y_max or 0), 1.0)
    min_px = max(int(min_px), 16)
    max_count = max(2, int(max(int(plot_height), min_px) / min_px))
    raw = y_max / max(max_count - 1, 1)
    magnitude = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
    step = 10 * magnitude
    for mult in (1, 2, 2.5, 5, 10):
        candidate = mult * magnitude
        if candidate + magnitude * 1e-9 >= raw:
            step = candidate
            break

    def pack(step_size):
        top = math.ceil((y_max / step_size) - 1e-9) * step_size
        values = [0.0]
        guard = 0
        while values[-1] + step_size <= top + step_size * 0.01 and guard < 40:
            values.append(values[-1] + step_size)
            guard += 1
        if values[-1] < y_max - step_size * 0.01:
            values.append(values[-1] + step_size)
        return values[-1], values

    top, values = pack(step)
    while len(values) > max_count:
        step *= 2
        top, values = pack(step)
    return top, values


def simulate_payoff(debts, settings, today=None, extras_by_period=None, minimum_only=False):
    """Interest and balances for minimums only, or for the payoff-order snowball.

    The snowball path follows the same payments as the estimated payoff dates.
    """
    today = today or datetime.date.today()
    extras_by_period = extras_by_period or {}
    ranked = [debt for debt in debts if _is_active_debt(debt)]
    ranked.sort(key=_payoff_rank_key)
    interest = {}
    dates = {}
    freq = settings.get("frequency", "biweekly")
    start = parse_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    if not ranked:
        return {
            "dates": dates,
            "interest": interest,
            "points": [(start, {})],
            "start": start,
            "frequency": freq,
        }

    ppy = periods_per_year(freq) or 12
    extra = 0.0 if minimum_only else additional_per_pay(debts, settings)
    display_periods = parse_period_count(settings.get("projection_periods")) or 120
    max_periods = max(display_periods, ppy * 50)
    rows = []
    balances = {}
    starts = {}
    debt_ids = []
    for debt in ranked:
        did = debt["id"]
        debt_ids.append(did)
        balances[did] = parse_amount(debt.get("current"), minimum=0, allow_blank=True) or 0.0
        interest[did] = 0.0
        started = parse_iso_date(debt.get("schedule_start"))
        if started:
            starts[did] = started
        rows.append((
            did,
            debt,
            parse_amount(debt.get("rate"), minimum=0, allow_blank=True) or 0.0,
            parse_amount(debt.get("min_payment"), minimum=0, allow_blank=True) or 0.0,
            annual_minimum(debt) / ppy,
        ))
    points = [(start, dict(balances))]
    current = start
    for _ in range(max_periods):
        nxt = add_period(current, freq)
        pay_text = current.isoformat()
        logged_amounts = {} if minimum_only else _logged_extra_amounts(extras_by_period, pay_text, debt_ids)
        use_logged = (not minimum_only) and (current <= today or bool(logged_amounts))
        after = {}
        for did, debt, rate, minimum, per_pay_min in rows:
            if did in dates:
                after[did] = 0.0
                continue
            bal = balances.get(did, 0.0)
            if starts.get(did) and current < starts[did]:
                after[did] = bal
                continue
            charged = bal * (rate / ppy) if rate > 0 else 0.0
            interest[did] = interest.get(did, 0.0) + charged
            after[did] = bal + charged - _period_minimum(debt, minimum, per_pay_min, current, nxt)
        if minimum_only or use_logged:
            for did, debt, rate, minimum, per_pay_min in rows:
                if did in dates or (starts.get(did) and current < starts[did]):
                    continue
                bal = after.get(did, 0.0)
                if bal < 0.5:
                    after[did] = 0.0
                    dates[did] = pay_text
                    if not minimum_only:
                        extra += per_pay_min if include_paid_minimum(debt) else 0.0
                    continue
                if minimum_only:
                    after[did] = bal
                    continue
                pay = min(logged_amounts.get(did, 0.0), bal)
                bal -= pay
                if bal < 0.5:
                    after[did] = 0.0
                    dates[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                else:
                    after[did] = bal
        else:
            pool = extra
            for did, debt, rate, minimum, per_pay_min in rows:
                if did in dates or (starts.get(did) and current < starts[did]):
                    continue
                bal = after.get(did, 0.0)
                if bal < 0.5:
                    after[did] = 0.0
                    dates[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                    continue
                if pool <= 0:
                    after[did] = bal
                    continue
                pay = min(pool, bal)
                bal -= pay
                pool -= pay
                if bal < 0.5:
                    after[did] = 0.0
                    dates[did] = pay_text
                    extra += per_pay_min if include_paid_minimum(debt) else 0.0
                else:
                    after[did] = bal
        balances = after
        points.append((current, dict(balances)))
        if len(dates) >= len(rows):
            break
        current = nxt
    for debt in ranked:
        dates.setdefault(debt["id"], "Never")
    return {
        "dates": dates,
        "interest": interest,
        "points": points,
        "start": start,
        "frequency": freq,
    }


def _active_debts_in_order(debts):
    rows = [debt for debt in debts if _is_active_debt(debt)]
    rows.sort(key=_payoff_rank_key)
    return rows


def chart_debt_lines(points, debts):
    """One balance line per debt that still has a balance."""
    active = _active_debts_in_order(debts)
    order = {debt["id"]: index for index, debt in enumerate(active)}
    series = []
    for debt in active:
        did = debt["id"]
        plotted = []
        for day, balances in points:
            total = max(0.0, balances.get(did, 0.0))
            plotted.append((day, total))
            if total < 0.5 and len(plotted) > 1:
                break
        if not plotted:
            continue
        series.append({
            "id": did,
            "name": debt.get("short_name") or f"Debt {did}",
            "color": CHART_COLORS[order.get(did, 0) % len(CHART_COLORS)],
            "balance": parse_amount(debt.get("current"), minimum=0, allow_blank=True) or 0.0,
            "points": plotted,
        })
    return series


def chart_type_lines(points, debts):
    """One balance line per debt type that still has a balance."""
    active = _active_debts_in_order(debts)
    by_type = {}
    for debt in active:
        by_type.setdefault(_debt_type_name(debt), []).append(debt)
    series = []
    for index, name in enumerate(OVERVIEW_TYPES):
        members = by_type.get(name)
        if not members:
            continue
        ids = [debt["id"] for debt in members]
        plotted = []
        for day, balances in points:
            total = sum(max(0.0, balances.get(did, 0.0)) for did in ids)
            plotted.append((day, total))
            if total < 0.5 and len(plotted) > 1:
                break
        if not plotted:
            continue
        balance = 0.0
        for debt in members:
            balance += parse_amount(debt.get("current"), minimum=0, allow_blank=True) or 0.0
        series.append({
            "id": name,
            "name": name,
            "color": CHART_COLORS[index % len(CHART_COLORS)],
            "balance": balance,
            "points": plotted,
        })
    return series


def smallest_chart_debt(series):
    """The line with the smallest balance."""
    if not series:
        return None
    return min(series, key=lambda item: (item["balance"], str(item["name"]).lower()))


def chart_time_scale(start, series):
    """Last day reached by the visible lines."""
    if start is None:
        return start, 1
    full_end = start
    for item in series or []:
        points = item.get("points") or []
        if points and points[-1][0] > full_end:
            full_end = points[-1][0]
    return full_end, max((full_end - start).days, 1)


def interest_totals(result, debt_ids):
    return sum(result["interest"].get(did, 0.0) for did in debt_ids)


def overview_summary(debts, settings, extras_by_period=None):
    """Income, minimums, and DTI for active debts (current balance of $0.50 or more)."""
    pay_frequency = settings.get("frequency", "biweekly")
    salary = parse_amount(settings.get("yearly_net_salary"), minimum=0, allow_blank=True) or 0.0
    spouse_freq = settings.get("spouse_income_frequency", "biweekly")
    other_freq = settings.get("other_income_frequency", "biweekly")
    spouse = parse_amount(settings.get("spouse_salary"), minimum=0, allow_blank=True) or 0.0
    other = parse_amount(settings.get("other_income"), minimum=0, allow_blank=True) or 0.0
    salary_pay = period_amount_for_income(salary, "yearly")
    spouse_pay = period_amount_for_income(spouse, spouse_freq)
    other_pay = period_amount_for_income(other, other_freq)
    yearly_total = salary + spouse + other
    monthly_income = yearly_total / 12.0
    pay_periods = periods_per_year(pay_frequency) or 12
    by_type = {name: {"remaining": 0.0, "annual_min": 0.0} for name in OVERVIEW_TYPES}
    has_mortgage = False
    mortgage_annual = 0.0
    for debt in debts:
        remaining = parse_amount(debt.get("current"), allow_blank=True)
        if remaining is None:
            continue
        if remaining < 0.5:
            continue
        kind = _debt_type_name(debt)
        annual = annual_minimum(debt)
        by_type[kind]["remaining"] += remaining
        by_type[kind]["annual_min"] += annual
        if kind == "Mortgage":
            has_mortgage = True
            mortgage_annual += annual
    dates, overall_free, by_id = debt_free_dates_by_type(
        debts, settings, extras_by_period=extras_by_period,
    )
    for name, row in by_type.items():
        row["debt_free"] = dates.get(name, "✓")
    by_debtor = []
    ordered = [debt for debt in debts if _is_active_debt(debt)]
    ordered.sort(key=_payoff_rank_key)
    for debt in ordered:
        remaining = parse_amount(debt.get("current"), allow_blank=True) or 0.0
        by_debtor.append({
            "name": debt.get("short_name") or f"Debt {debt.get('id')}",
            "remaining": remaining,
            "annual_min": annual_minimum(debt),
            "debt_free": by_id.get(debt.get("id"), "Never"),
        })
    remaining_total = sum(row["remaining"] for row in by_type.values())
    annual_total = sum(row["annual_min"] for row in by_type.values())
    monthly_min = annual_total / 12.0
    per_pay = annual_total / pay_periods if pay_periods else 0.0
    if monthly_income <= 0:
        front_dti = "—"
        back_dti = "—"
    else:
        back_dti = f"{(monthly_min / monthly_income) * 100:.1f}%"
        if has_mortgage:
            front_dti = f"{((mortgage_annual / 12.0) / monthly_income) * 100:.1f}%"
        else:
            front_dti = "N/A"
    return {
        "salary_yearly": salary,
        "spouse_yearly": spouse,
        "other_yearly": other,
        "salary_pay": salary_pay,
        "spouse_pay": spouse_pay,
        "other_pay": other_pay,
        "pay_income": None,
        "salary_monthly": salary / 12.0,
        "spouse_monthly": spouse / 12.0,
        "other_monthly": other / 12.0,
        "yearly_income": yearly_total,
        "monthly_income": monthly_income,
        "by_type": by_type,
        "by_debtor": by_debtor,
        "remaining_total": remaining_total,
        "monthly_min": monthly_min,
        "per_pay": per_pay,
        "pay_periods": pay_periods,
        "front_dti": front_dti,
        "back_dti": back_dti,
        "has_mortgage": has_mortgage,
        "debt_free_overall": overall_free,
    }


def month_length(year, month):
    if month == 2:
        leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        return 29 if leap else 28
    return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]


def parse_month_day(value):
    text = str(value or "").strip().replace("/", "-")
    if not text:
        return None
    parts = text.split("-")
    try:
        if len(parts) == 3:
            month, day = int(parts[1]), int(parts[2])
        elif len(parts) == 2:
            month, day = int(parts[0]), int(parts[1])
        else:
            return None
    except ValueError:
        return None
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return month, day


def next_occurrence(month, day, today=None):
    """Next MM-DD on or after today. December 20 + 01-21 lands in the following year."""
    today = today or datetime.date.today()

    def on(year):
        return datetime.date(year, month, min(day, month_length(year, month)))

    this_year = on(today.year)
    if this_year >= today:
        return this_year
    return on(today.year + 1)


def due_series(debt, today=None):
    parsed = parse_month_day(debt.get("due_md") or debt.get("due_date"))
    if not parsed:
        return None, "monthly"
    frequency = debt.get("due_frequency") or "monthly"
    if frequency not in DUE_FREQUENCIES:
        frequency = "monthly"
    anchor = None
    anchor_text = str(debt.get("due_anchor") or "")
    if len(anchor_text) >= 10:
        anchor = parse_flexible_date(anchor_text[:10])
    if anchor is None:
        anchor = next_occurrence(parsed[0], parsed[1], today)
    return anchor, frequency


def normalize_debt_due(debt, today=None):
    parsed = parse_month_day(debt.get("due_md") or debt.get("due_date"))
    frequency = debt.get("due_frequency") or "monthly"
    if frequency not in DUE_FREQUENCIES:
        frequency = "monthly"
    debt["due_frequency"] = frequency
    if not parsed:
        debt["due_md"] = ""
        debt["due_anchor"] = ""
        debt["due_date"] = ""
        return debt
    month, day = parsed
    debt["due_md"] = f"{month:02d}-{day:02d}"
    if not debt.get("due_anchor"):
        raw = str(debt.get("due_date") or "")
        if len(raw) >= 10:
            legacy = parse_flexible_date(raw[:10])
            if legacy is not None and legacy.month == month:
                debt["due_anchor"] = legacy.isoformat()
        if not debt.get("due_anchor"):
            debt["due_anchor"] = next_occurrence(month, day, today).isoformat()
    debt["due_date"] = debt["due_anchor"]
    if debt.get("debt_type") not in DEBT_TYPES:
        debt.pop("debt_type", None)
    return debt


def monthly_due_dates(due_day, period_start, next_period):
    found = []
    year = period_start.year
    month = period_start.month
    for _ in range(6):
        due = datetime.date(year, month, min(int(due_day), month_length(year, month)))
        if period_start <= due < next_period:
            found.append(due)
        if due >= next_period:
            break
        month += 1
        if month > 12:
            month = 1
            year += 1
    return found


def occurrences_between(anchor, frequency, period_start, next_period):
    """Due dates on or after period_start and before the next pay period."""
    if next_period <= period_start:
        return []
    if frequency == "monthly":
        return monthly_due_dates(anchor.day, period_start, next_period)
    step = DUE_INTERVAL_DAYS.get(frequency)
    if not step:
        return []
    cursor = anchor
    if cursor < period_start:
        jumps = (period_start - cursor).days // step
        cursor = cursor + timedelta(days=jumps * step)
        if cursor < period_start:
            cursor = cursor + timedelta(days=step)
    elif cursor >= next_period:
        jumps = ((cursor - next_period).days // step) + 1
        cursor = cursor - timedelta(days=jumps * step)
        while cursor < period_start:
            cursor = cursor + timedelta(days=step)
            if cursor >= next_period:
                return []
    found = []
    for _ in range(8):
        if cursor >= next_period:
            break
        if cursor >= period_start:
            found.append(cursor)
        cursor = cursor + timedelta(days=step)
    return found


def minimum_payments_due(debt, period_start, next_period):
    anchor, frequency = due_series(debt)
    if anchor is None:
        return 0
    return len(occurrences_between(anchor, frequency, period_start, next_period))


def _opening_before(next_opening, rate, minimum, extra, periods_year):
    payment = float(minimum) + float(extra or 0)
    periodic_rate = (float(rate) / periods_year) if rate and periods_year else 0.0
    if next_opening <= 0.005 and payment <= 0:
        return 0.0
    if periodic_rate:
        opening = (next_opening + payment) / (1 + periodic_rate)
    else:
        opening = next_opening + payment
    return round(max(0.0, opening), 2)


def calculate_schedule(debts, settings, extras_by_period, archives=None):
    freq = settings.get("frequency", "biweekly")
    ppy = max(1, periods_per_year(freq))
    n_periods = parse_period_count(settings.get("projection_periods")) or 120
    anchor = parse_date(settings.get("balance_as_of") or settings.get("first_pay_date"))
    display_start = schedule_start_date(settings, archives)

    known = {}
    rates = {}
    mins = {}
    usable = []
    for debt in debts:
        cleaned = sanitize_debt(debt)
        if cleaned is None or cleaned.get("schedule_hidden"):
            continue
        usable.append(cleaned)
        known[cleaned["id"]] = cleaned["current"]
        rates[cleaned["id"]] = cleaned["rate"]
        mins[cleaned["id"]] = cleaned["min_payment"]
    debts = usable
    archived = archived_pay_dates(archives)
    starts = {}
    for debt in debts:
        if debt.get("schedule_start"):
            starts[debt["id"]] = parse_date(debt["schedule_start"])
    historic = [debt for debt in debts if debt["id"] not in starts]
    historic_known = {debt["id"]: known[debt["id"]] for debt in historic}
    openings_at = {anchor: dict(historic_known)}

    cursor = anchor
    cursor_open = dict(historic_known)
    walked = 0
    while cursor > display_start and walked < 2600:
        walked += 1
        prev = subtract_period(cursor, freq)
        extras = mapping_get(extras_by_period, prev.isoformat(), {})
        if not isinstance(extras, dict):
            extras = {}
        previous_open = {}
        for debt in historic:
            did = debt["id"]
            per_pay_min = annual_minimum(debt) / ppy
            minimum_due = _period_minimum(debt, mins[did], per_pay_min, prev, cursor)
            previous_open[did] = _opening_before(
                cursor_open.get(did, 0.0),
                rates[did],
                minimum_due,
                parse_amount(mapping_get(extras, did, 0.0), minimum=0, allow_blank=True) or 0.0,
                ppy,
            )
        openings_at[prev] = previous_open
        cursor_open = previous_open
        cursor = prev

    schedule = []
    current = display_start
    openings = dict(openings_at.get(display_start, historic_known))
    produced = 0
    safety = 0
    while produced < n_periods and safety < n_periods + 2600:
        safety += 1
        nxt = add_period(current, freq)
        if nxt <= current:
            break
        if current in archived:
            current = nxt
            continue
        if current == anchor:
            for debt in historic:
                openings[debt["id"]] = known[debt["id"]]
        extras = mapping_get(extras_by_period, current.isoformat(), {})
        if not isinstance(extras, dict):
            extras = {}
        period_extras = {}
        new_balances = {}
        closings = {}
        total_extra = 0.0
        total_debt = 0.0
        estimating = current < anchor

        for debt in debts:
            did = debt["id"]
            extra = parse_amount(mapping_get(extras, did, 0.0), minimum=0, allow_blank=True) or 0.0
            start = starts.get(did)
            if start and current < start:
                new_balances[did] = None
                period_extras[did] = 0.0
                continue
            if start and current == start:
                bal = known[did]
            elif did not in starts and estimating:
                shown = openings_at.get(nxt, historic_known).get(did, 0.0)
                new_balances[did] = shown
                closings[did] = shown
                period_extras[did] = extra
                total_extra += extra
                total_debt += shown
                continue
            else:
                bal = openings.get(did, 0.0)

            if bal <= 0.005 and not (start and current == start and known[did] > 0.005):
                new_balances[did] = 0.0
                closings[did] = 0.0
                period_extras[did] = 0.0
                continue
            rate = rates[did]
            per_pay_min = annual_minimum(debt) / ppy
            minimum_due = _period_minimum(debt, mins[did], per_pay_min, current, nxt)
            interest = bal * (rate / ppy) if rate > 0 else 0.0
            ending = round(max(0.0, bal + interest - minimum_due - extra), 2)
            closings[did] = ending
            period_extras[did] = extra
            total_extra += extra
            shown = round(bal, 2) if (start and current == start) else ending
            new_balances[did] = shown
            total_debt += shown

        schedule.append({
            "index": len(schedule),
            "period": period_in_year(current, freq),
            "year": current.year,
            "date": current.isoformat(),
            "extras": period_extras,
            "balances": dict(new_balances),
            "closings": dict(closings),
            "total_extra": round(total_extra, 2),
            "total_debt": round(total_debt, 2),
        })

        still_ahead = any(start and current < start for start in starts.values())
        if current >= anchor:
            produced += 1
            if total_debt < 0.5 and not still_ahead:
                break
        openings = closings
        current = nxt

    return schedule


class SheetView(tk.Canvas):
    """Spreadsheet grid drawn with the same cells as the pay schedule."""

    def __init__(self, master, app):
        super().__init__(master, bg=COLORS["schedule_bg"], highlightthickness=0, bd=0)
        self._app = app
        self._role = "schedule"
        self._selected = None
        self._rows = []
        self._bounds = []
        self._header_bottom = 0
        self._fill_width = False
        self.inner = AppFrame(self)
        self.inner._role = "schedule"
        self.inner.configure(fg_color=COLORS["schedule_bg"])
        self._window = self.create_window((0, 0), window=self.inner, anchor="nw")
        tk.Misc.bind(self.inner, "<Configure>", self._sync_scroll)
        self.bind("<Configure>", self._stretch)

    def _sync_scroll(self, _event=None):
        box = self.bbox("all")
        if box:
            self.configure(scrollregion=box)
        try:
            self._fit_bar(getattr(self, "_hsb", None), self.inner.winfo_reqwidth(), self.winfo_width())
            self._fit_bar(getattr(self, "_vsb", None), self.inner.winfo_reqheight(), self.winfo_height())
        except tk.TclError:
            return

    def _fit_bar(self, bar, needed, visible):
        if bar is None or visible <= 1:
            return
        try:
            mapped = bool(bar.winfo_ismapped())
            if needed <= visible + 1:
                if mapped:
                    bar.grid_remove()
            elif not mapped:
                bar.grid()
        except tk.TclError:
            return

    def _stretch(self, event):
        """Keep columns at least as wide as their text. Grow them only when this sheet fills its pane."""
        self.update_idletasks()
        needed = max(self.inner.winfo_reqwidth(), 1)
        span = max(getattr(event, "width", 0) or 0, needed)
        if not self._fill_width:
            span = needed
        self.itemconfigure(self._window, width=span)
        self._sync_scroll()

    def _lock_column_floors(self):
        """Stop a filling sheet from shrinking a heading into the next column."""
        self.inner.update_idletasks()
        widest = {}
        for child in self.inner.winfo_children():
            info = child.grid_info()
            try:
                col = int(info.get("column", -1))
            except (TypeError, ValueError):
                continue
            if col < 0:
                continue
            widest[col] = max(widest.get(col, 0), child.winfo_reqwidth())
        for col, width in widest.items():
            current = int(self.inner.grid_columnconfigure(col).get("minsize") or 0)
            self.inner.grid_columnconfigure(col, minsize=max(current, width))

    def _measure(self):
        try:
            self.inner.update_idletasks()
            header = self.inner.grid_bbox(0, 0)
            self._header_bottom = (header[1] + header[3]) if header else 0
            bounds = []
            for index, row in enumerate(self._rows):
                box = self.inner.grid_bbox(0, index + 1)
                if box:
                    bounds.append((box[1], box[1] + box[3], row.get("iid")))
            self._bounds = bounds
        except tk.TclError:
            self._bounds = []

    def render(self, columns, rows, bind_row=None, fill_width=False, prominent_header=True):
        for child in self.inner.winfo_children():
            child.destroy()
        self._rows = []
        self._fill_width = bool(fill_width)
        for col, spec in enumerate(columns):
            anchor = spec.get("header_anchor", spec.get("anchor", "w"))
            label = self._app._make_sheet_header(self.inner, spec["title"], anchor, prominent=prominent_header)
            label.grid(row=0, column=col, sticky="nsew")
            floor = self._app._sheet_column_floor()
            weight = int(spec.get("weight") or 0) if self._fill_width else 0
            self.inner.grid_columnconfigure(
                col, minsize=max(floor, int(spec.get("minsize") or 0)), weight=weight,
            )
        self.inner.grid_rowconfigure(0, minsize=32)
        selected = None if self._selected is None else str(self._selected)
        for index, record in enumerate(rows):
            iid = record.get("iid")
            chosen = iid is not None and str(iid) == selected
            bg, fg, font, outline = self._app._sheet_row_style(index, record.get("kind", "body"), chosen)
            labels = []
            values = list(record.get("values") or [])
            for col, spec in enumerate(columns):
                text = values[col] if col < len(values) else ""
                label = self._app._make_sheet_cell(
                    self.inner, text, bg, fg, spec.get("anchor", "w"), font, outline,
                )
                label.grid(row=index + 1, column=col, sticky="nsew")
                label._sheet_iid = None if iid is None else str(iid)
                if bind_row and iid is not None:
                    bind_row(label)
                labels.append(label)
            self.inner.grid_rowconfigure(index + 1, minsize=28)
            self._rows.append({
                "iid": None if iid is None else str(iid),
                "kind": record.get("kind", "body"),
                "labels": labels,
            })
        if self._fill_width:
            self._lock_column_floors()
        self._stretch(type("Resize", (), {"width": self.winfo_width()})())
        self._measure()

    def _recolor(self):
        selected = None if self._selected is None else str(self._selected)
        for index, record in enumerate(self._rows):
            chosen = record.get("iid") is not None and record["iid"] == selected
            bg, fg, font, outline = self._app._sheet_row_style(index, record.get("kind", "body"), chosen)
            for label in record["labels"]:
                self._app._apply_sheet_cell(label, bg, fg, font, outline)
        self._measure()

    def selection(self):
        if self._selected is None:
            return ()
        return (self._selected,)

    def selection_set(self, item):
        self._selected = None if item in (None, "") else str(item)
        self._recolor()

    def see(self, item):
        self.update_idletasks()
        box = self.bbox("all")
        total = max(box[3], 1) if box else 1
        for y0, _y1, iid in self._bounds:
            if str(iid) == str(item):
                self.yview_moveto(max(0, y0) / total)
                return

    def get_children(self):
        return [row["iid"] for row in self._rows if row.get("iid") is not None]

    def identify_region(self, _x, y):
        if self.canvasy(y) < self._header_bottom:
            return "heading"
        if self.identify_row(y):
            return "cell"
        return "nothing"

    def identify_row(self, y):
        content_y = self.canvasy(y)
        for y0, y1, iid in self._bounds:
            if iid is None:
                continue
            if y0 <= content_y < y1:
                return iid
        return ""

    def index(self, item):
        for position, row in enumerate(self._rows):
            if row.get("iid") is not None and row["iid"] == str(item):
                return position
        return -1

    def move(self, item, _parent, index):
        current = self.index(item)
        if current < 0:
            return
        row = self._rows.pop(current)
        index = max(0, min(int(index), len(self._rows)))
        self._rows.insert(index, row)
        for visual, record in enumerate(self._rows, start=1):
            for label in record["labels"]:
                label.grid(row=visual)
        self._recolor()


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class DebtPayoffApp:
    def __init__(self, root):
        self.root = root
        apply_palette(self._peek_appearance())
        self.root.title("Debt Payoff Planner")
        self.root.minsize(1100, 650)
        self._place_restore_geometry()
        if isinstance(self.root, (ctk.CTk, ctk.CTkToplevel)):
            self.root.configure(fg_color=COLORS["bg"])
        else:
            self.root.configure(bg=COLORS["bg"])

        self.debts = []
        self.settings = deepcopy(DEFAULT_SETTINGS)
        self.extras_by_period = {}
        self.archives = {}
        self.current_file = None
        self.sample_view = False
        self.schedule = []
        self.undo_stack = []  # up to 15 snapshots
        self.redo_stack = []
        self._settings_lock = False
        self._display_scale = 1.0

        self._build_menu()
        self._build_ui()
        if not self._load_remembered_file():
            self.recalculate()
            self._mark_saved_state()
        self._set_title()
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_close)
        self._suspended_grab = None
        self.root.bind("<Unmap>", self._on_root_unmap, add="+")
        self.root.bind("<Map>", self._on_root_map, add="+")
        self.root._prepare_dialog = self._prepare_dialog
        self._present_window()

    def _place_restore_geometry(self):
        """Normal size used if the window is later restored from maximized."""
        scale = max(self._ui_scale(), 1.0)
        screen_w = max(self.root.winfo_screenwidth(), 1)
        screen_h = max(self.root.winfo_screenheight(), 1)
        width = min(1400, max(900, int((screen_w - 80) / scale)))
        height = min(800, max(600, int((screen_h - 120) / scale)))
        x = max(0, int((screen_w / scale - width) / 2))
        y = max(0, int((screen_h / scale - height) / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _present_window(self):
        """Show the window once, after the plan is loaded, maximized on this display."""
        self.root.withdraw()
        self.root._window_exists = True
        try:
            self.root.state("zoomed")
        except tk.TclError:
            self.root.deiconify()
            self._place_restore_geometry()
            return
        self._apply_titlebar()
        self.root.after(300, self._keep_maximized)
        self.root.after(300, self._apply_titlebar)
        self.root.after(1400, self._keep_maximized)
        self.root.after(1400, self._apply_titlebar)

    def _keep_maximized(self):
        try:
            if self.root.winfo_exists() and self.root.state() not in ("zoomed", "iconic", "withdrawn"):
                self.root.state("zoomed")
        except tk.TclError:
            return

    def _peek_appearance(self):
        try:
            with open(self._recent_file_path(), "r", encoding="utf-8") as handle:
                path = handle.read().strip()
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            mode = (data.get("settings") or {}).get("appearance")
        except Exception:
            mode = "dark"
        return "light" if mode == "light" else "dark"

    def _on_tab_changed(self):
        self._paint_tabs()
        self._align_snowball()
        self._refresh_visible()

    def _on_appearance(self, _event=None):
        if getattr(self, "_settings_lock", False):
            return
        mode = "dark" if self.appearance_var.get() == "Dark" else "light"
        if self.settings.get("appearance") == mode:
            return
        self.settings["appearance"] = mode
        self._apply_appearance(refresh=True)
        self.status_var.set("Appearance updated")

    def _apply_appearance(self, refresh=True):
        mode = apply_palette(self.settings.get("appearance"))
        self.settings["appearance"] = mode
        try:
            if isinstance(self.root, (ctk.CTk, ctk.CTkToplevel)):
                self.root.configure(fg_color=COLORS["bg"])
            else:
                self.root.configure(bg=COLORS["bg"])
        except tk.TclError:
            return
        self._recolor_widget(self.root)
        self._paint_menus()
        self._apply_titlebar()
        self._paint_tabs()
        if hasattr(self, "appearance_var"):
            locked = self._settings_lock
            self._settings_lock = True
            self.appearance_var.set("Dark" if mode == "dark" else "Light")
            self._settings_lock = locked
        if refresh:
            self._mark_views_dirty()
            self._refresh_visible()
            self._redraw_chart_surface()
        self._align_snowball()

    def _chrome_colors(self):
        """Title bar and File/Help strip. Dark uses the page color, which is darker than the cards."""
        if self.settings.get("appearance") == "dark":
            return COLORS["bg"], COLORS["header_fg"], COLORS["surface"]
        return COLORS["header_bg"], COLORS["header_fg"], COLORS["button_hover"]

    def _paint_menus(self):
        bg, fg, _hover = self._chrome_colors()
        bar = getattr(self, "_menu_bar_frame", None)
        if bar is not None:
            try:
                bar.configure(bg=bg)
            except tk.TclError:
                pass
        for button in (getattr(self, "_file_btn", None), getattr(self, "_help_btn", None)):
            if button is None:
                continue
            try:
                button.configure(bg=bg, fg=fg)
            except tk.TclError:
                pass
        for menu in (
            getattr(self, "_file_menu", None),
            getattr(self, "_help_menu", None),
        ):
            if menu is None:
                continue
            try:
                menu.configure(
                    bg=bg,
                    fg=fg,
                    activebackground=COLORS["button"],
                    activeforeground=COLORS["button_text"],
                )
            except tk.TclError:
                continue

    def _apply_titlebar(self, window=None):
        """Color a Windows caption to match File and Help, without hiding the window."""
        self._color_window_caption(window or self.root)
        if window is not None:
            return
        root = getattr(self, "root", None)
        if root is None:
            return
        try:
            children = root.winfo_children()
        except tk.TclError:
            return
        for child in children:
            if isinstance(child, tk.Toplevel):
                self._color_window_caption(child)

    def _color_window_caption(self, window):
        if not sys.platform.startswith("win") or window is None:
            return
        if getattr(window, "_caption_painting", False):
            return
        try:
            if not window.winfo_exists() or window.overrideredirect():
                return
            import ctypes
            window._caption_painting = True
            user32 = ctypes.windll.user32
            dwm = ctypes.windll.dwmapi
            user32.GetParent.argtypes = [ctypes.c_void_p]
            user32.GetParent.restype = ctypes.c_void_p
            dwm.DwmSetWindowAttribute.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint]
            dwm.DwmSetWindowAttribute.restype = ctypes.c_long
            user32.SetWindowPos.argtypes = [
                ctypes.c_void_p, ctypes.c_void_p,
                ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint,
            ]
            user32.SetWindowPos.restype = ctypes.c_bool
            hwnd = user32.GetParent(window.winfo_id())
            if not hwnd:
                hwnd = window.winfo_id()
            bg, _fg, _hover = self._chrome_colors()

            def colorref(hex_color):
                hex_color = hex_color.lstrip("#")
                red = int(hex_color[0:2], 16)
                green = int(hex_color[2:4], 16)
                blue = int(hex_color[4:6], 16)
                return ctypes.c_int(red | (green << 8) | (blue << 16))

            dark_buttons = ctypes.c_int(1 if self.settings.get("appearance") != "light" else 0)
            caption = colorref(bg)
            dwm.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(dark_buttons), ctypes.sizeof(dark_buttons))
            dwm.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(caption), ctypes.sizeof(caption))
            # The new caption stays invisible until Windows rebuilds the frame.
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0020 | 0x0040)
        except Exception:
            return
        finally:
            try:
                window._caption_painting = False
            except Exception:
                pass

    def _paint_tabs(self):
        notebook = getattr(self, "notebook", None)
        if notebook is None:
            return
        try:
            notebook.configure(
                fg_color=COLORS["bg"],
                border_color=COLORS["border"],
                segmented_button_fg_color=COLORS["tab_strip"],
                segmented_button_selected_color=COLORS["tab_active"],
                segmented_button_selected_hover_color=COLORS["tab_active_hover"],
                segmented_button_unselected_color=COLORS["tab_inactive"],
                segmented_button_unselected_hover_color=COLORS["tab_inactive_hover"],
                text_color=COLORS["tab_inactive_text"],
            )
        except tk.TclError:
            return
        try:
            notebook._bg_color = COLORS["bg"]
            notebook._configure_segmented_button_background_corners()
            tk.Frame.configure(notebook, bg=COLORS["bg"])
            notebook._canvas.configure(bg=COLORS["bg"])
        except tk.TclError:
            pass
        current = notebook.get()
        bar = getattr(notebook, "_segmented_button", None)
        if bar is None:
            return
        selected_h = 30
        rest_h = 26
        try:
            bar.configure(bg_color=COLORS["bg"], height=selected_h)
            bar.grid_rowconfigure(0, minsize=bar._apply_widget_scaling(selected_h))
            notebook.grid_rowconfigure(0, minsize=notebook._apply_widget_scaling(10 - (selected_h - rest_h)))
            notebook.grid_rowconfigure(1, minsize=notebook._apply_widget_scaling(8 + (selected_h - rest_h)))
        except tk.TclError:
            pass
        for name, button in bar._buttons_dict.items():
            try:
                active = name == current
                button.configure(
                    text_color=COLORS["tab_active_text"] if active else COLORS["tab_inactive_text"],
                    height=selected_h if active else rest_h,
                    font=ui_font( 13 if active else 12, "bold"),
                )
                if COLORS.get("tab_edge"):
                    button.configure(
                        border_width=0 if active else 1,
                        border_color=COLORS["tab_edge"],
                    )
                button.grid_configure(sticky="nsew" if active else "sew")
            except tk.TclError:
                continue
        try:
            notebook.update_idletasks()
        except tk.TclError:
            pass
        self._place_tab_accent()

    def _place_tab_accent(self):
        notebook = getattr(self, "notebook", None)
        if notebook is None:
            return
        line = getattr(self, "_tab_accent", None)
        if not COLORS.get("tab_line"):
            if line is not None:
                try:
                    line.place_forget()
                except tk.TclError:
                    pass
            return
        try:
            button = notebook._segmented_button._buttons_dict.get(notebook.get())
            if button is None:
                return
            if line is None or not line.winfo_exists():
                line = tk.Frame(notebook, bg=COLORS["tab_line"], height=3, highlightthickness=0, bd=0)
                self._tab_accent = line
            else:
                line.configure(bg=COLORS["tab_line"])
            x = button.winfo_rootx() - notebook.winfo_rootx()
            y = button.winfo_rooty() - notebook.winfo_rooty() + button.winfo_height() - 3
            line.place(x=x, y=max(0, y), width=max(button.winfo_width(), 1), height=3)
            line.lift()
        except tk.TclError:
            return

    def _redraw_chart_surface(self):
        canvas = getattr(self, "payoff_canvas", None)
        y_canvas = getattr(self, "chart_y_canvas", None)
        legend = getattr(self, "chart_legend_canvas", None)
        if canvas is None:
            return
        try:
            canvas.configure(bg=COLORS["chart_bg"], highlightbackground=COLORS["chart_frame"])
            if y_canvas is not None:
                y_canvas.configure(bg=COLORS["chart_bg"])
            if legend is not None:
                legend.configure(bg=COLORS["bg"])
        except tk.TclError:
            return
        if getattr(self, "_chart_model", None):
            self._draw_payoff_chart(reset_scroll=False)

    def _paint_combo(self, widget):
        """Interest dropdowns sit on a card, so they use the next lighter fill."""
        if getattr(widget, "_combo_fill", "") == "raised":
            fill = COLORS["button_secondary"]
            edge = COLORS["button_secondary_hover"]
        else:
            fill = COLORS["field"]
            edge = COLORS["border"]
        widget.configure(
            fg_color=fill,
            border_color=edge,
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            dropdown_fg_color=fill,
            dropdown_hover_color=COLORS["dropdown_hover"],
            dropdown_text_color=COLORS["text"],
            text_color=COLORS["text"],
        )

    def _recolor_widget(self, widget):
        try:
            if not widget.winfo_exists():
                return
        except tk.TclError:
            return
        role = getattr(widget, "_role", "")
        try:
            if role == "clear":
                pass
            elif role == "page":
                widget.configure(fg_color=COLORS["bg"])
            elif role == "card":
                widget.configure(fg_color=COLORS["surface"], border_color=COLORS["border"])
            elif role == "section_band":
                widget.configure(fg_color=_section_title_chrome()[0])
            elif role == "section":
                widget.configure(text_color=_section_title_chrome()[1])
            elif role == "section_note":
                widget.configure(text_color=_section_note_color())
            elif role == "section_rule":
                widget.configure(fg_color=_section_title_chrome()[2])
            elif isinstance(widget, AppLabel):
                if role == "status":
                    widget.configure(fg_color=COLORS["header_bg"], text_color=COLORS["header_fg"])
                elif role in _TITLE_ROLES:
                    widget.configure(text_color=COLORS["section_title"])
                elif role in _MUTED_ROLES:
                    widget.configure(text_color=COLORS["text_muted"])
                else:
                    widget.configure(text_color=COLORS["text"])
            elif role == "schedule":
                if isinstance(widget, tk.Canvas):
                    widget.configure(bg=COLORS["schedule_bg"])
                else:
                    widget.configure(fg_color=COLORS["schedule_bg"])
            elif role == "chart":
                widget.configure(bg=COLORS["chart_bg"])
                if widget is getattr(self, "payoff_canvas", None):
                    widget.configure(highlightbackground=COLORS["chart_frame"])
            elif role in ("legend", "archive"):
                widget.configure(bg=COLORS["bg"])
            elif role == "snow":
                widget.configure(bg=COLORS["snow_bg"])
            elif role == "snow_label":
                widget.configure(bg=COLORS["snow_bg"], fg=COLORS["snow_label"])
            elif role == "snow_value":
                widget.configure(bg=COLORS["snow_bg"], fg=COLORS["snow_value"])
            elif isinstance(widget, AppButton):
                tone = getattr(widget, "_button_tone", "primary")
                if tone == "secondary":
                    widget.configure(
                        fg_color=COLORS["button_secondary"],
                        hover_color=COLORS["button_secondary_hover"],
                        text_color=COLORS["button_secondary_text"],
                        border_color=COLORS["button_secondary_border"],
                    )
                elif tone == "danger":
                    widget.configure(
                        fg_color=COLORS["danger"],
                        hover_color=COLORS["danger_hover"],
                        text_color=COLORS["button_text"],
                        border_color=COLORS["danger"],
                    )
                else:
                    widget.configure(
                        fg_color=COLORS["button"],
                        hover_color=COLORS["button_hover"],
                        text_color=COLORS["button_text"],
                    )
            elif isinstance(widget, ctk.CTkEntry):
                widget.configure(
                    fg_color=COLORS["field"],
                    border_color=COLORS["border"],
                    text_color=COLORS["text"],
                )
            elif isinstance(widget, ctk.CTkComboBox):
                self._paint_combo(widget)
            elif isinstance(widget, ctk.CTkCheckBox):
                widget.configure(
                    fg_color=COLORS["button"],
                    hover_color=COLORS["button_hover"],
                    text_color=COLORS["text"],
                    checkmark_color=COLORS["button_text"],
                )
            elif isinstance(widget, ctk.CTkScrollbar):
                widget.configure(
                    button_color=COLORS["border"],
                    button_hover_color=COLORS["button"],
                    fg_color=COLORS["bg"],
                )
            elif isinstance(widget, ctk.CTkToplevel):
                widget.configure(fg_color=COLORS["bg"])
        except tk.TclError:
            pass
        try:
            children = widget.winfo_children()
        except tk.TclError:
            return
        for child in children:
            self._recolor_widget(child)

    def _ui_scale(self):
        """Match the old app, which Windows stretched on a high-DPI screen."""
        try:
            from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker
            scale = float(ScalingTracker.get_window_dpi_scaling(self.root))
        except Exception:
            scale = 1.0
        if scale < 1:
            scale = 1.0
        self._display_scale = scale
        return scale

    def _tk_font(self, size, *styles):
        scaled = max(1, int(round(float(size) * self._ui_scale())))
        if styles:
            return ui_font( scaled, *styles)
        return ui_font( scaled)

    def _data_family(self):
        """Small-text face for grids of numbers. Windows prefers Segoe UI Variable."""
        cached = getattr(self, "_data_family_name", None)
        if cached:
            return cached
        family = UI_FONT_FAMILY
        if sys.platform.startswith("win"):
            try:
                if "Segoe UI Variable Text" in set(tkfont.families(self.root)):
                    family = "Segoe UI Variable Text"
            except tk.TclError:
                family = "Segoe UI"
        self._data_family_name = family
        return family

    def _data_font(self, size, *styles):
        scaled = max(1, int(round(float(size) * self._ui_scale())))
        if styles:
            return (self._data_family(), scaled, *styles)
        return (self._data_family(), scaled)

    def _grid_font(self, size, weight=None):
        spec = self._data_font(size, weight) if weight else self._data_font(size)
        cache = self.__dict__.setdefault("_grid_fonts", {})
        font = cache.get(spec)
        if font is None:
            font = tkfont.Font(font=spec)
            cache[spec] = font
        return font

    def _mark_views_dirty(self):
        self._views_dirty = {"Pay Schedule", "Debts", "Overview", "Charts", "Interest", "Archives"}

    def _refresh_visible(self):
        """Paint the open tab. Hidden tabs wait until the user selects them."""
        dirty = getattr(self, "_views_dirty", None)
        if not dirty or not hasattr(self, "notebook"):
            return
        try:
            name = self.notebook.get()
        except tk.TclError:
            return
        if name == "Pay Schedule" and "Pay Schedule" in dirty:
            self.refresh_schedule_tree()
            dirty.discard("Pay Schedule")
        elif name == "Debts" and "Debts" in dirty:
            self.refresh_debt_tree()
            dirty.discard("Debts")
        elif name == "Overview" and "Overview" in dirty:
            self.refresh_overview()
            dirty.discard("Overview")
        elif name in ("Charts", "Interest") and ("Charts" in dirty or "Interest" in dirty):
            self.refresh_charts()
            dirty.discard("Charts")
            dirty.discard("Interest")
        elif name == "Archives" and "Archives" in dirty:
            self.refresh_archives()
            dirty.discard("Archives")

    def _fit_archive_canvas(self, _event=None):
        size = getattr(self, "_archive_size", None)
        canvas = getattr(self, "archive_canvas", None)
        if not size or canvas is None:
            return
        try:
            canvas.configure(scrollregion=(0, 0, size[0], size[1]))
        except tk.TclError:
            return

    def _prepare_dialog(self, dialog):
        """Keep a dialog with the main window and use the same caption colors."""
        try:
            dialog._deactivate_windows_window_header_manipulation = True
        except Exception:
            pass
        dialog.transient(self.root)
        apply_window_icon(dialog)
        self._color_window_caption(dialog)

        def repaint(event=None, target=dialog):
            if event is not None and event.widget is not target:
                return
            self._color_window_caption(target)

        try:
            dialog.bind("<Map>", repaint, add="+")
        except tk.TclError:
            pass
        for delay in (50, 300):
            dialog.after(delay, lambda target=dialog: self._color_window_caption(target))

    def _on_root_unmap(self, event):
        if event.widget is not self.root:
            return
        self.root.after(1, self._suspend_grab_if_iconified)

    def _suspend_grab_if_iconified(self):
        try:
            if not self.root.winfo_exists() or self.root.state() != "iconic":
                return
            grab = self.root.grab_current()
            if grab is not None:
                self._suspended_grab = grab
                grab.grab_release()
            for child in self.root.winfo_children():
                if isinstance(child, tk.Toplevel):
                    child.withdraw()
        except tk.TclError:
            self._suspended_grab = None

    def _on_root_map(self, event):
        if event.widget is not self.root:
            return
        self.root.after(1, self._restore_dialogs_after_minimize)

    def _root_is_open(self):
        """True while the main window is visible, including maximized."""
        try:
            return self.root.winfo_exists() and self.root.state() in ("normal", "zoomed")
        except tk.TclError:
            return False

    def _restore_dialogs_after_minimize(self):
        try:
            if not self._root_is_open():
                return
        except tk.TclError:
            return
        grab = self._suspended_grab
        self._suspended_grab = None
        windows = []
        if grab is not None:
            windows.append(grab)
        for child in self.root.winfo_children():
            if isinstance(child, tk.Toplevel) and child not in windows:
                windows.append(child)
        for window in windows:
            try:
                if window.winfo_exists():
                    window.deiconify()
                    window.lift()
            except tk.TclError:
                continue
        if grab is not None:
            try:
                if grab.winfo_exists():
                    grab.grab_set()
                    grab.focus_force()
            except tk.TclError:
                pass
        self._align_snowball()

    def _build_menu(self):
        bg, fg, hover = self._chrome_colors()
        file_menu = tk.Menu(
            self.root, tearoff=0, bg=bg, fg=fg,
            activebackground=COLORS["button"], activeforeground=COLORS["button_text"],
        )
        file_menu.add_command(label="New Blank", command=self.new_blank)
        file_menu.add_separator()
        file_menu.add_command(label="Open...", command=self.open_file)
        file_menu.add_command(label="Save", command=self.save_file)
        file_menu.add_command(label="Save As...", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Export Schedule to Excel...", command=self.export_schedule)
        file_menu.add_command(label="Export Archives to Excel...", command=self.export_archives)
        file_menu.add_separator()
        file_menu.add_command(label="Undo", command=self.undo)
        file_menu.add_command(label="Redo", command=self.redo)
        file_menu.add_separator()
        file_menu.add_command(label="Check for Updates", command=lambda: self.check_for_updates(interactive=True))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.confirm_close)

        help_menu = tk.Menu(
            self.root, tearoff=0, bg=bg, fg=fg,
            activebackground=COLORS["button"], activeforeground=COLORS["button_text"],
        )
        help_menu.add_command(label="How to Use", command=self.show_help)
        menu_font = self._tk_font(10)
        for menu in (file_menu, help_menu):
            try:
                menu.configure(font=menu_font)
            except tk.TclError:
                pass
        self._file_menu = file_menu
        self._help_menu = help_menu

        bar = tk.Frame(self.root, bg=bg, highlightthickness=0, bd=0)
        bar.pack(fill="x", side="top")
        self._menu_bar_frame = bar
        self._file_btn = self._menu_heading(bar, "File", file_menu)
        self._help_btn = self._menu_heading(bar, "Help", help_menu)
        self._file_btn.pack(side="left")
        self._help_btn.pack(side="left")
        try:
            self.root.configure(menu="")
        except tk.TclError:
            pass

    def _menu_heading(self, parent, text, menu):
        bg, fg, hover = self._chrome_colors()
        pad_x = max(14, int(round(12 * self._ui_scale())))
        pad_y = max(6, int(round(5 * self._ui_scale())))
        label = tk.Label(
            parent, text=text, bg=bg, fg=fg, font=self._tk_font(10),
            padx=pad_x, pady=pad_y, cursor="hand2",
        )

        def on_enter(_event):
            _bg, _fg, hot = self._chrome_colors()
            label.configure(bg=hot)

        def on_leave(_event):
            label.configure(bg=self._chrome_colors()[0])

        def on_click(_event):
            _bg, _fg, hot = self._chrome_colors()
            label.configure(bg=hot)
            menu.tk_popup(label.winfo_rootx(), label.winfo_rooty() + label.winfo_height())

        label.bind("<Enter>", on_enter)
        label.bind("<Leave>", on_leave)
        label.bind("<Button-1>", on_click)
        return label

    def _build_ui(self):
        notebook = ctk.CTkTabview(
            self.root,
            fg_color=COLORS["bg"],
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=8,
            segmented_button_fg_color=COLORS["tab_strip"],
            segmented_button_selected_color=COLORS["tab_active"],
            segmented_button_selected_hover_color=COLORS["tab_active_hover"],
            segmented_button_unselected_color=COLORS["tab_inactive"],
            segmented_button_unselected_hover_color=COLORS["tab_inactive_hover"],
            text_color=COLORS["tab_active_text"],
            segmented_button_font=ui_font( 12, "bold"),
            anchor="w",
            command=self._on_tab_changed,
        )
        notebook.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        self.notebook = notebook
        select_tab = notebook.set

        def _select_tab(name):
            select_tab(name)
            self._paint_tabs()
            self._refresh_visible()

        notebook.set = _select_tab
        for name, builder in (
            ("Overview", self._build_overview_tab),
            ("Pay Schedule", self._build_schedule_tab),
            ("Debts", self._build_debts_tab),
            ("Charts", self._build_charts_tab),
            ("Interest", self._build_interest_tab),
            ("Archives", self._build_archives_tab),
            ("Settings", self._build_settings_tab),
        ):
            builder(notebook.add(name))

        self.snowball_var = tk.StringVar(value="$0.00")
        snowball = tk.Toplevel(self.root)
        snowball.overrideredirect(True)
        snowball.configure(bg=COLORS["snow_bg"])
        snowball.withdraw()
        try:
            snowball.transient(self.root)
            snowball.attributes("-toolwindow", True)
        except tk.TclError:
            pass
        snowball._role = "snow"
        pad = self._ui_scale()
        snow_label = tk.Label(
            snowball, text="Additional Per Pay", bg=COLORS["snow_bg"], fg=COLORS["snow_label"],
            font=self._tk_font(11, "bold"),
        )
        snow_label._role = "snow_label"
        snow_label.pack(side="left", padx=(int(round(10 * pad)), int(round(8 * pad))))
        snow_value = tk.Label(
            snowball, textvariable=self.snowball_var, bg=COLORS["snow_bg"], fg=COLORS["snow_value"],
            font=self._tk_font(13, "bold"),
        )
        snow_value._role = "snow_value"
        snow_value.pack(side="left", padx=(0, int(round(10 * pad))))
        self._snowball = snowball
        tk.Misc.bind(notebook, "<Configure>", self._align_snowball, add="+")
        self.root.bind("<Configure>", self._align_snowball, add="+")
        self.root.after_idle(self._align_snowball)

        self.status_var = tk.StringVar(value="Ready")
        status = AppLabel(self.root, textvariable=self.status_var, anchor="w", role="status", background=COLORS["header_bg"], foreground=COLORS["header_fg"])
        status.pack(side="bottom", fill="x")
        self._status = status
        self._paint_tabs()

    def _selected_tab_box(self):
        notebook = self.notebook
        try:
            name = notebook.get()
            button = notebook._segmented_button._buttons_dict.get(name)
        except (AttributeError, tk.TclError):
            return None
        if button is None:
            return None
        try:
            if not button.winfo_exists():
                return None
            top = button.winfo_rooty() - notebook.winfo_rooty()
            height = button.winfo_height()
        except tk.TclError:
            return None
        if height < 8:
            return None
        return top, height

    def _align_snowball(self, _event=None):
        if getattr(self, "_aligning_snowball", False):
            return
        snowball = getattr(self, "_snowball", None)
        notebook = getattr(self, "notebook", None)
        if snowball is None or notebook is None:
            return
        self._aligning_snowball = True
        try:
            if not self._root_is_open():
                return
            if not notebook.winfo_exists() or not notebook.get():
                return
            notebook.update_idletasks()
            box = self._selected_tab_box()
            if box is None:
                return
            top, height = box
            if height < 8:
                return
            snowball.update_idletasks()
            natural = max(snowball.winfo_reqheight(), 1)
            if natural > height:
                top -= (natural - height) // 2
                height = natural
            width = max(snowball.winfo_reqwidth(), 1)
            x = notebook.winfo_rootx() + notebook.winfo_width() - width - 4
            y = notebook.winfo_rooty() + top
            snowball.geometry(f"{width}x{height}+{x}+{y}")
            snowball.deiconify()
            if self.root.grab_current() is None:
                try:
                    snowball.attributes("-topmost", True)
                    snowball.lift()
                    snowball.attributes("-topmost", False)
                except tk.TclError:
                    snowball.lift()
            self._place_tab_accent()
        except tk.TclError:
            return
        finally:
            self._aligning_snowball = False

    def _debt_columns(self):
        return (
            {"title": "Short Name", "anchor": "w", "minsize": 110, "weight": 2},
            {"title": "Full Name", "anchor": "w", "minsize": 190, "weight": 3},
            {"title": "Current Balance", "anchor": "e", "minsize": 140, "weight": 2},
            {"title": "Rate", "anchor": "e", "minsize": 80, "weight": 1},
            {"title": "Minimum Payment", "anchor": "e", "minsize": 150, "weight": 2},
            {"title": "Next Due", "anchor": "center", "header_anchor": "center", "minsize": 120, "weight": 2},
            {"title": "Payment Frequency", "anchor": "center", "header_anchor": "center", "minsize": 160, "weight": 2},
            {"title": "Debt Type", "anchor": "center", "header_anchor": "center", "minsize": 120, "weight": 2},
            {"title": "Payoff Rank", "anchor": "center", "header_anchor": "center", "minsize": 110, "weight": 1},
            {"title": "Estimated Payoff", "anchor": "center", "header_anchor": "center", "minsize": 150, "weight": 2},
        )

    def _bind_debt_cell(self, label):
        label.bind("<ButtonPress-1>", self._on_debt_drag_press, add="+")
        label.bind("<B1-Motion>", self._on_debt_drag_motion, add="+")
        label.bind("<ButtonRelease-1>", self._on_debt_drag_release, add="+")

    def _build_debts_tab(self, parent):
        sheet_host = AppFrame(parent)
        sheet_host.pack(fill="both", expand=True, padx=8, pady=8)
        self.debt_tree = self._pack_sheet(sheet_host)
        self._debt_drag = None

        btn_frame = AppFrame(parent)
        btn_frame.pack(fill="x", padx=8, pady=6)
        self.debt_btn_frame = btn_frame
        AppButton(btn_frame, text="Add Debt", command=self.add_debt).pack(side="left", padx=3)
        AppButton(btn_frame, text="Edit Debt", command=self.edit_debt).pack(side="left", padx=3)
        AppButton(btn_frame, text="Delete Debt", style="danger", command=self.delete_debt).pack(side="left", padx=3)
        AppButton(btn_frame, text="Edit Debt Balances", command=self.update_balances).pack(side="left", padx=3)
        AppButton(btn_frame, text="Move Down", style="secondary", command=lambda: self.move_debt(1)).pack(side="right", padx=3)
        AppButton(btn_frame, text="Move Up", style="secondary", command=lambda: self.move_debt(-1)).pack(side="right", padx=3)

        self.refresh_debt_tree()

    def _build_settings_tab(self, parent):
        bar = AppFrame(parent)
        bar.pack(fill="x", padx=20, pady=(16, 0))
        AppLabel(bar, text="Appearance", style="Header.TLabel").pack(side="left")
        self.appearance_var = tk.StringVar(value="Dark" if self.settings.get("appearance") == "dark" else "Light")
        appearance = AppCombo(
            bar, textvariable=self.appearance_var, values=("Light", "Dark"), state="readonly", width=12,
        )
        appearance.pack(side="left", padx=(12, 0))
        appearance.bind("<<ComboboxSelected>>", self._on_appearance)
        self.settings_appearance = appearance

        frm = AppFrame(parent, padding=20)
        frm.pack(fill="both", expand=True)

        self._settings_label(frm, 0, "First Pay Date (MM-DD-YYYY):")
        self.date_var = tk.StringVar(value=self.settings["first_pay_date"])
        date_entry = AppEntry(frm, textvariable=self.date_var, width=24)
        date_entry.grid(row=0, column=1, sticky="w", pady=6)
        date_entry.bind("<FocusOut>", self._commit_settings)
        date_entry.bind("<Return>", self._commit_settings)
        self.settings_first_pay = date_entry

        self._settings_label(frm, 1, "Balances As Of (MM-DD-YYYY):")
        self.anchor_date_var = tk.StringVar(value=self.settings.get("balance_as_of") or self.settings["first_pay_date"])
        anchor_entry = AppEntry(frm, textvariable=self.anchor_date_var, width=24)
        anchor_entry.grid(row=1, column=1, sticky="w", pady=6)
        anchor_entry.bind("<FocusOut>", self._commit_settings)
        anchor_entry.bind("<Return>", self._commit_settings)
        self.settings_balance_as_of = anchor_entry

        self._settings_label(frm, 2, "Starting Additional Per Pay:")
        self.starting_var = tk.StringVar(value=format_money_entry(self.settings.get("starting_additional", 0)))
        starting_entry = AppEntry(frm, textvariable=self.starting_var, width=18)
        starting_entry.grid(row=2, column=1, sticky="w", pady=6)
        starting_entry.bind("<FocusOut>", self._commit_settings)
        starting_entry.bind("<Return>", self._commit_settings)
        self.settings_starting = starting_entry

        AppLabel(frm, text="Projection Periods:", style="Header.TLabel").grid(row=3, column=0, sticky="w", pady=6)
        self.proj_var = tk.StringVar(value=str(self.settings["projection_periods"]))
        proj_entry = AppEntry(frm, textvariable=self.proj_var, width=24)
        proj_entry.grid(row=3, column=1, sticky="w", pady=6)
        proj_entry.bind("<FocusOut>", self._commit_settings)
        proj_entry.bind("<Return>", self._commit_settings)
        self.settings_proj = proj_entry

        self._settings_label(frm, 4, "Yearly Net Salary:")
        salary_row = AppFrame(frm)
        salary_row.grid(row=4, column=1, sticky="w", pady=6)
        self.salary_var = tk.StringVar(value=format_money_entry(self.settings.get("yearly_net_salary", 0)))
        salary_entry = AppEntry(salary_row, textvariable=self.salary_var, width=18)
        salary_entry.pack(side="left")
        salary_entry.bind("<FocusOut>", self._commit_settings)
        salary_entry.bind("<Return>", self._commit_settings)
        self.settings_salary = salary_entry
        AppLabel(salary_row, text="Pay Frequency").pack(side="left", padx=(12, 0))
        self.freq_var = tk.StringVar(value=self.settings["frequency"])
        freq_combo = AppCombo(
            salary_row, textvariable=self.freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        freq_combo.pack(side="left", padx=(8, 0))
        freq_combo.bind("<<ComboboxSelected>>", self._commit_settings)
        self.settings_freq = freq_combo

        self._settings_label(frm, 5, "Spouse / Secondary Yearly Net Income:")
        spouse_row = AppFrame(frm)
        spouse_row.grid(row=5, column=1, sticky="w", pady=6)
        self.spouse_var = tk.StringVar(value=format_money_entry(self.settings.get("spouse_salary", 0)))
        spouse_entry = AppEntry(spouse_row, textvariable=self.spouse_var, width=18)
        spouse_entry.pack(side="left")
        spouse_entry.bind("<FocusOut>", self._commit_settings)
        spouse_entry.bind("<Return>", self._commit_settings)
        self.settings_spouse = spouse_entry
        AppLabel(spouse_row, text="Pay Frequency").pack(side="left", padx=(12, 0))
        spouse_freq_value = self.settings.get("spouse_income_frequency", "biweekly")
        if spouse_freq_value not in PAY_FREQUENCIES:
            spouse_freq_value = "biweekly"
        self.spouse_freq_var = tk.StringVar(value=spouse_freq_value)
        spouse_freq = AppCombo(
            spouse_row, textvariable=self.spouse_freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        spouse_freq.pack(side="left", padx=(8, 0))
        spouse_freq.bind("<<ComboboxSelected>>", self._commit_settings)
        self.settings_spouse_freq = spouse_freq

        self._settings_label(frm, 6, "Other Yearly Net Salary:")
        other_row = AppFrame(frm)
        other_row.grid(row=6, column=1, sticky="w", pady=6)
        self.other_income_var = tk.StringVar(value=format_money_entry(self.settings.get("other_income", 0)))
        other_entry = AppEntry(other_row, textvariable=self.other_income_var, width=18)
        other_entry.pack(side="left")
        other_entry.bind("<FocusOut>", self._commit_settings)
        other_entry.bind("<Return>", self._commit_settings)
        self.settings_other = other_entry
        AppLabel(other_row, text="Pay Frequency").pack(side="left", padx=(12, 0))
        other_freq_value = self.settings.get("other_income_frequency", "biweekly")
        if other_freq_value not in PAY_FREQUENCIES:
            other_freq_value = "biweekly"
        self.other_freq_var = tk.StringVar(value=other_freq_value)
        other_freq = AppCombo(
            other_row, textvariable=self.other_freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        other_freq.pack(side="left", padx=(8, 0))
        other_freq.bind("<<ComboboxSelected>>", self._commit_settings)
        self.settings_other_freq = other_freq

    def _build_schedule_tab(self, parent):
        # Frozen Period/Pay Date plus a separately scrolling money grid
        container = AppFrame(parent)
        container.pack(fill="both", expand=True, padx=8, pady=8)

        self.sched_frozen_canvas = tk.Canvas(
            container, bg=COLORS["schedule_bg"], highlightthickness=0, width=int(round(170 * self._ui_scale())),
        )
        self.sched_canvas = tk.Canvas(container, bg=COLORS["schedule_bg"], highlightthickness=0)
        vsb = AppScroll(container, orient="vertical")
        hsb = AppScroll(container, orient="horizontal", command=self.sched_canvas.xview)

        def _sched_yview(*args):
            self.sched_frozen_canvas.yview(*args)
            self.sched_canvas.yview(*args)
            self._redraw_visible_rows(self.sched_frozen_canvas)
            self._redraw_visible_rows(self.sched_canvas)

        def _sched_yscroll(first, last):
            vsb.set(first, last)

        vsb.configure(command=_sched_yview)
        self.sched_frozen_canvas.configure(yscrollcommand=_sched_yscroll)
        self.sched_canvas.configure(yscrollcommand=_sched_yscroll, xscrollcommand=hsb.set)

        self.sched_frozen_canvas.grid(row=0, column=0, sticky="nsw")
        self.sched_canvas.grid(row=0, column=1, sticky="nsew")
        vsb.grid(row=0, column=2, sticky="ns")
        hsb.grid(row=1, column=1, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        self.sched_frozen_inner = AppFrame(self.sched_frozen_canvas)
        self.sched_inner = AppFrame(self.sched_canvas)
        self.sched_frozen_inner._role = "schedule"
        self.sched_inner._role = "schedule"
        self.sched_frozen_canvas._role = "schedule"
        self.sched_canvas._role = "schedule"
        self.sched_frozen_inner.configure(fg_color=COLORS["schedule_bg"])
        self.sched_inner.configure(fg_color=COLORS["schedule_bg"])
        self.sched_frozen_canvas.create_window((0, 0), window=self.sched_frozen_inner, anchor="nw")
        self.sched_canvas.create_window((0, 0), window=self.sched_inner, anchor="nw")

        def _fit_schedule_panes(_event=None):
            frozen = getattr(self, "_sched_frozen_size", None)
            body = getattr(self, "_sched_body_size", None)
            if frozen:
                self.sched_frozen_canvas.configure(
                    scrollregion=(0, 0, frozen[0], frozen[1]), width=max(frozen[0], 1),
                )
            if body:
                self.sched_canvas.configure(scrollregion=(0, 0, body[0], body[1]))

        tk.Misc.bind(self.sched_frozen_inner, "<Configure>", _fit_schedule_panes)
        tk.Misc.bind(self.sched_inner, "<Configure>", _fit_schedule_panes)
        self._fit_schedule_panes = _fit_schedule_panes

        def _on_mousewheel(event):
            x, y = self.root.winfo_pointerxy()
            widget = self.root.winfo_containing(x, y)
            schedule_parts = (
                self.sched_canvas, self.sched_inner,
                self.sched_frozen_canvas, self.sched_frozen_inner,
            )
            while widget is not None:
                if widget in (getattr(self, "archive_canvas", None), getattr(self, "archive_inner", None)):
                    self._archive_yview("scroll", int(-1 * (event.delta / 120)), "units")
                    return
                for name in ("debt_tree", "overview_debtor_sheet", "overview_type_sheet"):
                    sheet = getattr(self, name, None)
                    if sheet is not None and widget in (sheet, getattr(sheet, "inner", None)):
                        sheet.yview_scroll(int(-1 * (event.delta / 120)), "units")
                        return
                if widget in schedule_parts:
                    _sched_yview("scroll", int(-1 * (event.delta / 120)), "units")
                    return
                widget = widget.master

        self.sched_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        btn_frame = AppFrame(parent)
        btn_frame.pack(fill="x", padx=8, pady=6)
        AppLabel(btn_frame, text="Double-click a row to assign extras for that period", style="OverviewCaption.TLabel").pack(side="left")
        archive_btn = AppButton(btn_frame, text="Archive Completed Pay Periods", command=self.archive_years)
        archive_btn.pack(side="right", padx=3)
        self.sched_archive_btn = archive_btn
        self.sched_btn_frame = btn_frame

    def _build_archives_tab(self, parent):
        year_bar = AppFrame(parent)
        year_bar.pack(fill="x", padx=8, pady=(10, 0))
        self.archive_year_bar = year_bar
        self._archive_year = None

        container = AppFrame(parent)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        self.archive_canvas = tk.Canvas(container, bg=COLORS["schedule_bg"], highlightthickness=0)
        self.archive_canvas._role = "schedule"
        vsb = AppScroll(container, orient="vertical", command=self._archive_yview)
        hsb = AppScroll(container, orient="horizontal", command=self.archive_canvas.xview)
        self.archive_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.archive_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.archive_inner = AppFrame(self.archive_canvas)
        self.archive_inner._role = "schedule"
        self.archive_inner.configure(fg_color=COLORS["schedule_bg"])
        self.archive_canvas.create_window((0, 0), window=self.archive_inner, anchor="nw")
        tk.Misc.bind(self.archive_inner, "<Configure>", self._fit_archive_canvas)
        self.refresh_archives()

    def _build_overview_tab(self, parent):
        frm = AppFrame(parent, padding=16)
        frm.pack(fill="both", expand=True)

        cards = AppFrame(frm)
        cards.grid(row=0, column=0, columnspan=2, sticky="nsew")
        cards.grid_columnconfigure(0, weight=1)
        cards.grid_columnconfigure(1, weight=1)
        cards.grid_columnconfigure(2, weight=1)

        income = AppSection(cards, text="Income", style="Overview.TLabelframe", padding=10)
        income.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        value_pad = (16, 0)
        AppLabel(income, text="Yearly", style="OverviewValue.TLabel", anchor="e").grid(row=0, column=1, sticky="ew", padx=value_pad)
        AppLabel(
            income, text="Pay Period Average", style="OverviewValue.TLabel",
            anchor="e", justify="right", wraplength=140,
        ).grid(row=0, column=2, sticky="ew", padx=value_pad)
        AppLabel(
            income, text="Monthly Average", style="OverviewValue.TLabel",
            anchor="e", justify="right", wraplength=140,
        ).grid(row=0, column=3, sticky="ew", padx=value_pad)
        self.overview_salary_year = tk.StringVar(value="$0.00")
        self.overview_salary_pay = tk.StringVar(value="$0.00")
        self.overview_salary_month = tk.StringVar(value="$0.00")
        self.overview_spouse_year = tk.StringVar(value="$0.00")
        self.overview_spouse_pay = tk.StringVar(value="$0.00")
        self.overview_spouse_month = tk.StringVar(value="$0.00")
        self.overview_other_year = tk.StringVar(value="$0.00")
        self.overview_other_pay = tk.StringVar(value="$0.00")
        self.overview_other_month = tk.StringVar(value="$0.00")
        self.overview_yearly_income = tk.StringVar(value="$0.00")
        self.overview_pay_income = tk.StringVar(value="$0.00")
        self.overview_monthly_income = tk.StringVar(value="$0.00")
        AppLabel(income, text="Yearly Net Salary", style="OverviewValue.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        AppLabel(income, textvariable=self.overview_salary_year, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=1, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_salary_pay, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=2, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_salary_month, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=3, sticky="ew", padx=value_pad)
        AppLabel(income, text="Spouse / Secondary Yearly Net Income", style="OverviewValue.TLabel").grid(row=2, column=0, sticky="w", pady=3)
        AppLabel(income, textvariable=self.overview_spouse_year, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=1, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_spouse_pay, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=2, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_spouse_month, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=3, sticky="ew", padx=value_pad)
        AppLabel(income, text="Other Yearly Net Salary", style="OverviewValue.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        AppLabel(income, textvariable=self.overview_other_year, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=1, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_other_pay, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=2, sticky="ew", padx=value_pad)
        AppLabel(income, textvariable=self.overview_other_month, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=3, sticky="ew", padx=value_pad)
        AppLabel(income, text="Total", style="OverviewValue.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 2))
        AppLabel(income, textvariable=self.overview_yearly_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=1, sticky="ew", padx=value_pad, pady=(12, 2))
        AppLabel(income, textvariable=self.overview_pay_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=2, sticky="ew", padx=value_pad, pady=(12, 2))
        AppLabel(income, textvariable=self.overview_monthly_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=3, sticky="ew", padx=value_pad, pady=(12, 2))
        income.grid_columnconfigure(0, weight=1)
        for column in (1, 2, 3):
            income.grid_columnconfigure(column, weight=1, uniform="income_value")
        self.overview_income = income

        dti = AppSection(cards, text="Debt-to-Income", style="Overview.TLabelframe", padding=12)
        dti.grid(row=0, column=1, sticky="nsew", padx=4)
        self.overview_front_dti = tk.StringVar(value="—")
        self.overview_back_dti = tk.StringVar(value="—")
        AppLabel(dti, text="Front-End", style="OverviewValue.TLabel").grid(row=0, column=0, sticky="w")
        AppLabel(dti, text="Mortgage Only", style="OverviewCaption.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 2))
        AppLabel(dti, textvariable=self.overview_front_dti, style="OverviewHero.TLabel").grid(row=0, column=1, rowspan=2, sticky="e", padx=8)
        AppLabel(dti, text="Back-End", style="OverviewValue.TLabel").grid(row=2, column=0, sticky="w", pady=(14, 0))
        AppLabel(dti, text="All Active Debts", style="OverviewCaption.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 2))
        AppLabel(dti, textvariable=self.overview_back_dti, style="OverviewHero.TLabel").grid(row=2, column=1, rowspan=2, sticky="e", padx=8, pady=(14, 0))
        dti.grid_columnconfigure(0, weight=1)
        self.overview_dti = dti

        mins = AppSection(cards, text="Minimum Payments", style="Overview.TLabelframe", padding=12)
        mins.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        self.overview_monthly_min = tk.StringVar(value="$0.00")
        self.overview_per_pay = tk.StringVar(value="$0.00")
        self.overview_per_pay_label = tk.StringVar(value="Per Pay Period")
        AppLabel(mins, text="Monthly", style="OverviewValue.TLabel").grid(row=0, column=0, sticky="w", pady=10)
        AppLabel(mins, textvariable=self.overview_monthly_min, style="OverviewHero.TLabel").grid(row=0, column=1, sticky="e", pady=10)
        AppLabel(mins, textvariable=self.overview_per_pay_label, style="OverviewValue.TLabel").grid(row=1, column=0, sticky="w", pady=10)
        AppLabel(mins, textvariable=self.overview_per_pay, style="OverviewHero.TLabel").grid(row=1, column=1, sticky="e", pady=10)
        mins.grid_columnconfigure(0, weight=1)
        self.overview_mins = mins

        tables = AppFrame(frm)
        tables.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        tables.grid_columnconfigure(0, weight=1, uniform="overview_tables")
        tables.grid_columnconfigure(1, weight=1, uniform="overview_tables")
        tables.grid_rowconfigure(0, weight=1)

        debtor = AppSection(
            tables,
            text="Debt by Debtor",
            note="Payoff Date with Estimated Snowball Payments",
            style="Overview.TLabelframe",
            padding=10,
        )
        debtor.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.overview_debtor_sheet = self._pack_sheet(debtor)
        self.overview_debtor = debtor

        types = AppSection(tables, text="Debt by Type", style="Overview.TLabelframe", padding=10)
        types.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.overview_type_sheet = self._pack_sheet(types)
        self.overview_types = types
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(1, weight=1)
        self.refresh_overview()

    def _overview_sheet_columns(self, name_title):
        center = {"header_anchor": "center"}
        return (
            {"title": name_title, "anchor": "w", "minsize": 160, "weight": 2},
            {"title": "Remaining", "anchor": "e", "minsize": 130, "weight": 1},
            {"title": "Monthly Minimum", "anchor": "e", "minsize": 160, "weight": 1},
            {"title": "Per Pay Period", "anchor": "e", "minsize": 150, "weight": 1},
            {"title": "Payoff Date", "anchor": "center", "minsize": 150, "weight": 2, **center},
        )

    def refresh_overview(self):
        if not hasattr(self, "overview_debtor_sheet"):
            return
        summary = overview_summary(self.debts, self.settings, self.extras_by_period)
        money = lambda value: f"${value:,.2f}"
        self.overview_salary_year.set(money(summary["salary_yearly"]))
        self.overview_salary_pay.set(money(summary["salary_pay"]))
        self.overview_salary_month.set(money(summary["salary_monthly"]))
        self.overview_spouse_year.set(money(summary["spouse_yearly"]))
        self.overview_spouse_pay.set(money(summary["spouse_pay"]))
        self.overview_spouse_month.set(money(summary["spouse_monthly"]))
        self.overview_other_year.set(money(summary["other_yearly"]))
        self.overview_other_pay.set(money(summary["other_pay"]))
        self.overview_other_month.set(money(summary["other_monthly"]))
        self.overview_yearly_income.set(money(summary["yearly_income"]))
        self.overview_pay_income.set("—" if summary["pay_income"] is None else money(summary["pay_income"]))
        self.overview_monthly_income.set(money(summary["monthly_income"]))
        self.overview_front_dti.set(summary["front_dti"])
        self.overview_back_dti.set(summary["back_dti"])
        self.overview_monthly_min.set(money(summary["monthly_min"]))
        frequency = self.settings.get("frequency", "biweekly")
        freq_label = {
            "biweekly": "Bi-Weekly",
            "semimonthly": "Semi-Monthly",
            "monthly": "Monthly",
        }.get(frequency, "Bi-Weekly")
        self.overview_per_pay_label.set(freq_label)
        self.overview_per_pay.set(money(summary["per_pay"]))
        pay_periods = summary["pay_periods"] or 12

        def sheet_rows(entries):
            rows = []
            for entry in entries:
                rows.append({
                    "kind": "body",
                    "values": (
                        entry["name"],
                        money(entry["remaining"]),
                        money(entry["annual_min"] / 12.0),
                        money(entry["annual_min"] / pay_periods),
                        format_display_date(entry.get("debt_free") or "✓"),
                    ),
                })
            rows.append({
                "kind": "total",
                "values": (
                    "Total",
                    money(summary["remaining_total"]),
                    money(summary["monthly_min"]),
                    money(summary["per_pay"]),
                    format_display_date(summary.get("debt_free_overall") or "✓"),
                ),
            })
            return rows

        debtor_rows = []
        for row in summary.get("by_debtor") or []:
            debtor_rows.append(row)
        self.overview_debtor_sheet.render(
            self._overview_sheet_columns("Debtor"), sheet_rows(debtor_rows), fill_width=True, prominent_header=True,
        )
        type_rows = []
        for name in OVERVIEW_TYPES:
            entry = dict(summary["by_type"][name])
            entry["name"] = name
            type_rows.append(entry)
        self.overview_type_sheet.render(
            self._overview_sheet_columns("Debt Type"), sheet_rows(type_rows), fill_width=True, prominent_header=True,
        )

    # ----- Debt management -----
    def _debt_total_now(self, debt):
        """Balance still owed on the payday now in effect, after that day's extra."""
        stored = parse_amount(debt.get("current"), allow_blank=True) or 0.0
        schedule = getattr(self, "schedule", None) or []
        if not schedule:
            return stored
        anchor = parse_date(self.settings.get("balance_as_of") or self.settings.get("first_pay_date"))
        active = current_pay_date(self.settings)
        if anchor is None or active < anchor:
            return stored
        period = next((row for row in schedule if row.get("date") == active.isoformat()), None)
        if not period:
            return stored
        closing = mapping_get(period.get("closings") or {}, debt.get("id"), None)
        if closing is None:
            return stored
        value = parse_amount(closing, minimum=0, allow_blank=True)
        return stored if value is None else value

    def refresh_debt_tree(self):
        selected_id = None
        sel = self.debt_tree.selection()
        if sel:
            try:
                selected_id = int(sel[0])
            except (TypeError, ValueError):
                selected_id = None
        self._debts_by_id = {}
        for debt in self.debts:
            try:
                self._debts_by_id[int(debt["id"])] = debt
            except (TypeError, ValueError):
                continue
        ranked = assign_payoff_ranks(self.debts)
        _dates, _overall, payoff_by_id = debt_free_dates_by_type(
            self.debts, self.settings, extras_by_period=self.extras_by_period,
        )
        stored_paid = [debt for debt in self.debts if not _is_active_debt(debt)]
        stored_paid.sort(key=lambda debt: int(debt.get("id") or 0))
        open_debts = []
        paid = []
        for debt in ranked + stored_paid:
            if self._debt_total_now(debt) < 0.5:
                paid.append(debt)
            else:
                open_debts.append(debt)
        rows = []
        for d in open_debts + paid:
            is_paid = self._debt_total_now(d) < 0.5
            if is_paid:
                rank_text = "—"
                est_text = "✓"
            else:
                rank_text = str(d.get("payoff_rank") or "")
                est_text = format_display_date(payoff_by_id.get(d.get("id"), "Never"))
            current = self._debt_total_now(d)
            rate = parse_amount(d.get("rate"), allow_blank=True) or 0.0
            minimum = parse_amount(d.get("min_payment"), allow_blank=True) or 0.0
            due_anchor, _due_frequency = due_series(d)
            due_text = format_display_date(due_anchor) if due_anchor else "—"
            rows.append({
                "iid": str(d["id"]),
                "kind": "paid" if is_paid else "body",
                "values": (
                    d.get("short_name") or f"Debt {d.get('id')}",
                    d.get("full_name") or "",
                    f"${current:,.2f}",
                    f"{rate * 100:.2f}%",
                    f"${minimum:,.2f}",
                    due_text,
                    DUE_FREQUENCY_LABELS.get(d.get("due_frequency"), "—") if d.get("due_md") else "—",
                    d.get("debt_type") if d.get("debt_type") in DEBT_TYPES else "—",
                    rank_text,
                    est_text,
                ),
            })
        self.debt_tree._selected = str(selected_id) if selected_id is not None else None
        self.debt_tree.render(self._debt_columns(), rows, bind_row=self._bind_debt_cell)
        if selected_id is not None and str(selected_id) in self.debt_tree.get_children():
            self.debt_tree.see(str(selected_id))

    def _debt_by_iid(self, item):
        try:
            did = int(item)
        except (TypeError, ValueError):
            return None
        found = getattr(self, "_debts_by_id", {}).get(did)
        if found is not None:
            return found
        return next((debt for debt in self.debts if debt.get("id") == did), None)

    def _commit_payoff_order(self, ordered):
        if self._block_sample_edit():
            return
        self.push_undo()
        assign_payoff_ranks(self.debts, ordered)
        self.refresh_debt_tree()
        self.refresh_overview()
        self.refresh_schedule_tree()

    def move_debt(self, direction):
        sel = self.debt_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select a debt to move.")
            return
        debt = self._debt_by_iid(sel[0])
        if debt is None or self._debt_total_now(debt) < 0.5:
            self.status_var.set("Only unpaid debts can be ranked")
            return
        ordered = assign_payoff_ranks(self.debts)
        try:
            index = ordered.index(debt)
        except ValueError:
            return
        new_index = index + int(direction)
        if new_index < 0 or new_index >= len(ordered):
            return
        ordered[index], ordered[new_index] = ordered[new_index], ordered[index]
        self._commit_payoff_order(ordered)

    def _debt_pointer_y(self):
        return self.root.winfo_pointery() - self.debt_tree.winfo_rooty()

    def _on_debt_drag_press(self, event):
        y = self._debt_pointer_y()
        if self.debt_tree.identify_region(event.x, y) not in ("cell", "tree"):
            self._debt_drag = None
            return
        item = self.debt_tree.identify_row(y)
        debt = self._debt_by_iid(item) if item else None
        if debt is None:
            self._debt_drag = None
            return
        self.debt_tree.selection_set(item)
        if self._debt_total_now(debt) < 0.5 or getattr(self, "sample_view", False):
            self._debt_drag = None
            return
        self._debt_drag = {"item": item, "moved": False}

    def _on_debt_drag_motion(self, event):
        drag = self._debt_drag
        if not drag:
            return
        target = self.debt_tree.identify_row(self._debt_pointer_y())
        if not target or target == drag["item"]:
            return
        target_debt = self._debt_by_iid(target)
        if target_debt is None or self._debt_total_now(target_debt) < 0.5:
            return
        self.debt_tree.move(drag["item"], "", self.debt_tree.index(target))
        drag["moved"] = True
        self.debt_tree.configure(cursor="fleur")

    def _on_debt_drag_release(self, event):
        drag = self._debt_drag
        self._debt_drag = None
        try:
            self.debt_tree.configure(cursor="")
        except tk.TclError:
            pass
        if not drag or not drag.get("moved"):
            return
        ordered = []
        for item in self.debt_tree.get_children():
            debt = self._debt_by_iid(item)
            if debt is None or self._debt_total_now(debt) < 0.5:
                continue
            ordered.append(debt)
        self._commit_payoff_order(ordered)

    def _unique_short_name(self, name, skip_id=None):
        base = (name or "Debt").strip() or "Debt"
        taken = {d.get("short_name") for d in self.debts if d.get("id") != skip_id}
        if base not in taken:
            return base
        n = 2
        while f"{base} ({n})" in taken:
            n += 1
        return f"{base} ({n})"

    def add_debt(self):
        if self._block_sample_edit():
            return
        dialog = DebtDialog(self.root, title="Add Debt")
        if dialog.result:
            self.push_undo()
            new_id = max((int(d["id"]) for d in self.debts), default=0) + 1
            dialog.result["id"] = new_id
            dialog.result["short_name"] = self._unique_short_name(dialog.result.get("short_name"))
            dialog.result["schedule_start"] = current_pay_date(self.settings).isoformat()
            dialog.result.pop("payoff_rank", None)
            self.debts.append(dialog.result)
            self.refresh_debt_tree()
            self.recalculate()

    def edit_debt(self):
        if self._block_sample_edit():
            return
        sel = self.debt_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select a debt to edit.")
            return
        debt = self._debt_by_iid(sel[0])
        if not debt:
            messagebox.showerror("Error", "Could not find selected debt.")
            return
        dialog = DebtDialog(self.root, title="Edit Debt", debt=debt)
        if dialog.result:
            self.push_undo()
            dialog.result["short_name"] = self._unique_short_name(dialog.result.get("short_name"), skip_id=debt.get("id"))
            debt.update(dialog.result)
            self._sync_paid_off_state(debt)
            self.refresh_debt_tree()
            self.recalculate()

    def delete_debt(self):
        if self._block_sample_edit():
            return
        sel = self.debt_tree.selection()
        if not sel:
            return
        debt = self._debt_by_iid(sel[0])
        if not debt:
            return
        short = debt.get("short_name") or "this debt"
        minimum = parse_amount(debt.get("min_payment"), minimum=0, allow_blank=True) or 0.0
        if include_paid_minimum(debt):
            effect = (
                f"If you delete this debt, its minimum payment (${minimum:,.2f}) will no longer "
                "be used in Additional Per Pay calculations."
            )
        else:
            effect = (
                f"This debt is set to leave its minimum payment (${minimum:,.2f}) out of "
                "Additional Per Pay. Deleting it removes the debt from the plan."
            )
        warning = f"Delete {short}?\n\n{effect}\n\nDo you want to continue deleting this debt?"
        if messagebox.askyesno("Delete Debt", warning, default=messagebox.NO):
            self._freeze_archive_columns()
            self.push_undo()
            self.debts = [d for d in self.debts if d.get("id") != debt.get("id")]
            self.refresh_debt_tree()
            self.recalculate()
            self.refresh_archives()

    def update_balances(self):
        if self._block_sample_edit():
            return
        if not self.debts:
            return
        win = ctk.CTkToplevel(self.root)
        self._prepare_dialog(win)
        win.title("Edit Debt Balances")
        count = max(len(self.debts), 1)
        scale = max(self._ui_scale(), 1.0)
        screen_h = max(win.winfo_screenheight(), 600)
        height = min(int(screen_h / scale * 0.82), 180 + count * 46)
        win.geometry(f"560x{max(height, 360)}")
        win.configure(fg_color=COLORS["bg"])
        holder = tk.Frame(win, bg=COLORS["bg"])
        canvas = tk.Canvas(holder, bg=COLORS["bg"], highlightthickness=0, bd=0)
        scroll = AppScroll(holder, command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        inner = tk.Frame(canvas, bg=COLORS["bg"])
        inner.grid_columnconfigure(1, weight=1)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def fit_list(event=None):
            canvas.configure(scrollregion=canvas.bbox("all") or (0, 0, 0, 0))
            width = canvas.winfo_width()
            if width > 1:
                canvas.itemconfigure(window_id, width=width)
            needed = inner.winfo_reqheight()
            if needed > canvas.winfo_height() + 2:
                if not scroll.winfo_ismapped():
                    scroll.pack(side="right", fill="y")
            elif scroll.winfo_ismapped():
                scroll.pack_forget()

        inner.bind("<Configure>", fit_list)
        canvas.bind("<Configure>", fit_list)

        def on_wheel(event):
            if inner.winfo_reqheight() > canvas.winfo_height() + 2:
                canvas.yview_scroll(int(-event.delta / 120), "units")
            return "break"

        def bind_wheel(widget):
            try:
                widget.bind("<MouseWheel>", on_wheel, add="+")
            except Exception:
                pass
            for child in widget.winfo_children():
                bind_wheel(child)

        canvas.bind("<MouseWheel>", on_wheel)
        canvas.pack(side="left", fill="both", expand=True)
        entries = {}
        for i, d in enumerate(self.debts):
            top = 8 if i == 0 else 6
            AppLabel(inner, text=d["short_name"]).grid(row=i, column=0, sticky="w", padx=(8, 10), pady=(top, 6))
            var = tk.StringVar(value=str(d["current"]))
            ent = AppEntry(inner, textvariable=var, width=16)
            ent.grid(row=i, column=1, sticky="ew", padx=(0, 8), pady=(top, 6))
            entries[d["id"]] = var
            ent.bind("<FocusOut>", lambda _event: commit(), add="+")
            ent.bind("<Return>", lambda _event: finish(), add="+")
        bind_wheel(inner)

        def commit():
            changes = []
            for debt in self.debts:
                text = entries[debt["id"]].get().strip()
                value = parse_amount(text, minimum=0, maximum=1e12, allow_blank=True)
                if value is None:
                    self.status_var.set(f"Enter a non-negative number for {debt.get('short_name') or 'that debt'}")
                    return False
                current = parse_amount(debt.get("current"), allow_blank=True) or 0.0
                if abs(value - current) > 0.0001:
                    changes.append((debt, value))
            if changes:
                self.push_undo()
                for debt, value in changes:
                    debt["current"] = value
                    self._sync_paid_off_state(debt)
                self.refresh_debt_tree()
                self.recalculate()
            return True

        def finish():
            if commit():
                win.destroy()

        footer = AppFrame(win)
        footer.pack(side="bottom", fill="x", pady=(4, 18))
        AppLabel(footer, text="Press Done to save these balances.", style="OverviewCaption.TLabel").pack(pady=(4, 8))
        AppButton(footer, text="Done", command=finish).pack(pady=(0, 2))
        holder.pack(fill="both", expand=True, padx=8, pady=(16, 8))
        win.protocol("WM_DELETE_WINDOW", win.destroy)

    # ----- Settings -----
    def _settings_label(self, parent, row, text):
        AppLabel(parent, text=text, style="Header.TLabel").grid(row=row, column=0, sticky="w", pady=6, padx=(0, 16))

    def _format_valid_money_fields(self):
        self._settings_lock = True
        try:
            for var in (self.starting_var, self.salary_var, self.spouse_var, self.other_income_var):
                formatted = format_money_entry(var.get())
                if formatted is not None:
                    var.set(formatted)
        finally:
            self._settings_lock = False

    def _sync_settings_form(self):
        self._settings_lock = True
        try:
            self.freq_var.set(self.settings.get("frequency", "biweekly"))
            self.date_var.set(format_display_date(self.settings.get("first_pay_date", "")))
            if hasattr(self, "anchor_date_var"):
                self.anchor_date_var.set(format_display_date(self.settings.get("balance_as_of") or self.settings.get("first_pay_date", "")))
            if hasattr(self, "starting_var"):
                self.starting_var.set(format_money_entry(self.settings.get("starting_additional", 0)))
            self.proj_var.set(str(self.settings.get("projection_periods", 120)))
            if hasattr(self, "salary_var"):
                self.salary_var.set(format_money_entry(self.settings.get("yearly_net_salary", 0)))
            if hasattr(self, "spouse_var"):
                self.spouse_var.set(format_money_entry(self.settings.get("spouse_salary", 0)))
            if hasattr(self, "spouse_freq_var"):
                freq = self.settings.get("spouse_income_frequency", "biweekly")
                self.spouse_freq_var.set(freq if freq in PAY_FREQUENCIES else "biweekly")
            if hasattr(self, "other_income_var"):
                self.other_income_var.set(format_money_entry(self.settings.get("other_income", 0)))
            if hasattr(self, "other_freq_var"):
                freq = self.settings.get("other_income_frequency", "biweekly")
                self.other_freq_var.set(freq if freq in PAY_FREQUENCIES else "biweekly")
            if hasattr(self, "appearance_var"):
                self.appearance_var.set("Dark" if self.settings.get("appearance") == "dark" else "Light")
        finally:
            self._settings_lock = False

    def _commit_settings(self, _event=None):
        if getattr(self, "_settings_lock", False):
            return
        if self._block_sample_edit():
            return
        self._format_valid_money_fields()
        frequency = self.freq_var.get()
        if frequency not in ("biweekly", "semimonthly", "monthly"):
            self.status_var.set("Choose a pay frequency")
            return
        first = parse_iso_date(self.date_var.get())
        balance = parse_iso_date(self.anchor_date_var.get())
        if first is None or balance is None:
            self.status_var.set("Use MM-DD-YYYY for dates")
            return
        periods = parse_period_count(self.proj_var.get())
        if periods is None:
            self.status_var.set("Projection periods must be a whole number from 1 to 2600")
            return
        starting = parse_amount(self.starting_var.get(), minimum=0, maximum=1e12, allow_blank=True)
        salary = parse_amount(self.salary_var.get(), minimum=0, maximum=1e12, allow_blank=True)
        spouse = parse_amount(self.spouse_var.get(), minimum=0, maximum=1e12, allow_blank=True)
        other_income = parse_amount(self.other_income_var.get(), minimum=0, maximum=1e12, allow_blank=True)
        if None in (starting, salary, spouse, other_income):
            self.status_var.set("Use non-negative numbers for additional payment and income")
            return
        spouse_freq = income_frequency_key(self.spouse_freq_var.get() if hasattr(self, "spouse_freq_var") else "biweekly")
        other_freq = income_frequency_key(self.other_freq_var.get() if hasattr(self, "other_freq_var") else "biweekly")
        if spouse_freq not in PAY_FREQUENCIES:
            spouse_freq = "biweekly"
        if other_freq not in PAY_FREQUENCIES:
            other_freq = "biweekly"
        proposed = {
            "frequency": frequency,
            "first_pay_date": first.isoformat(),
            "balance_as_of": balance.isoformat(),
            "projection_periods": periods,
            "starting_additional": starting,
            "yearly_net_salary": salary,
            "spouse_salary": spouse,
            "other_income": other_income,
            "spouse_income_frequency": spouse_freq,
            "other_income_frequency": other_freq,
            "income_entered_as_yearly": True,
        }
        current = {key: self.settings.get(key) for key in proposed}
        current["projection_periods"] = parse_period_count(current.get("projection_periods"))
        for key in ("starting_additional", "yearly_net_salary", "spouse_salary", "other_income"):
            current[key] = parse_amount(current.get(key), minimum=0, allow_blank=True) or 0.0
        if proposed == current:
            if self.status_var.get() in (
                "Choose a pay frequency",
                "Use MM-DD-YYYY for dates",
                "Projection periods must be a whole number from 1 to 2600",
                "Use non-negative numbers for additional payment and income",
            ):
                self.status_var.set("Settings updated")
            return
        self.push_undo()
        self.settings.update(proposed)
        self.recalculate()
        live_start = parse_iso_date(self.settings.get("first_pay_date"))
        if live_start and first < live_start:
            self.status_var.set("First Pay Date cannot be before archived pay periods")
        else:
            self.status_var.set("Settings updated")

    # ----- Schedule -----
    def _history_snapshot(self):
        return {
            "debts": deepcopy(self.debts),
            "extras_by_period": deepcopy(self.extras_by_period),
            "settings": deepcopy(self.settings),
            "archives": deepcopy(self.archives),
        }

    def push_undo(self):
        snap = self._history_snapshot()
        self.undo_stack.append(snap)
        if len(self.undo_stack) > 15:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self):
        if self._block_sample_edit():
            return
        if not self.undo_stack:
            self.status_var.set("Nothing to undo")
            return
        current = self._history_snapshot()
        self.redo_stack.append(current)
        snap = self.undo_stack.pop()
        self.debts = snap["debts"]
        self.extras_by_period = snap["extras_by_period"]
        self.settings = snap["settings"]
        self.archives = snap.get("archives", {})
        self._apply_appearance(refresh=False)
        self._sync_settings_form()
        self.refresh_debt_tree()
        self.refresh_archives()
        self.recalculate()
        self.status_var.set(f"Undo ({len(self.undo_stack)} left)")

    def redo(self):
        if self._block_sample_edit():
            return
        if not self.redo_stack:
            self.status_var.set("Nothing to redo")
            return
        current = self._history_snapshot()
        self.undo_stack.append(current)
        if len(self.undo_stack) > 15:
            self.undo_stack.pop(0)
        snap = self.redo_stack.pop()
        self.debts = snap["debts"]
        self.extras_by_period = snap["extras_by_period"]
        self.settings = snap["settings"]
        self.archives = snap.get("archives", {})
        self._apply_appearance(refresh=False)
        self._sync_settings_form()
        self.refresh_debt_tree()
        self.refresh_archives()
        self.recalculate()
        self.status_var.set(f"Redo ({len(self.redo_stack)} left)")

    def _migrate_extras_to_dates(self):
        """Keep extras attached to a pay date when earlier periods are added."""
        if not self.extras_by_period:
            return
        if all(isinstance(key, str) and len(key) == 10 and key[4] == "-" and key[7] == "-" for key in self.extras_by_period):
            return
        anchor = parse_date(self.settings.get("balance_as_of") or self.settings.get("first_pay_date"))
        freq = self.settings.get("frequency", "biweekly")
        count = (parse_period_count(self.settings.get("projection_periods")) or 120) + 5
        dates = []
        current = anchor
        for _ in range(count):
            dates.append(current.isoformat())
            current = add_period(current, freq)
        migrated = {}
        for key, value in self.extras_by_period.items():
            key_text = str(key)
            if len(key_text) == 10 and key_text[4] == "-" and key_text[7] == "-":
                target = key_text
            else:
                try:
                    index = int(key)
                except (TypeError, ValueError):
                    continue
                if index < 0 or index >= len(dates):
                    continue
                target = dates[index]
            if isinstance(value, dict):
                migrated[target] = normalize_id_map(value)
        self.extras_by_period = migrated

    def _clip_live_schedule_after_archives(self):
        """Keep the live grid after the last archived payday."""
        last = last_archived_pay_date(self.archives)
        if last is None:
            return
        freq = self.settings.get("frequency", "biweekly")
        anchor = parse_date(self.settings.get("balance_as_of") or self.settings.get("first_pay_date"))
        nxt = align_on_or_after(last + timedelta(days=1), anchor, freq)
        self.settings["earliest_pay_date"] = nxt.isoformat()
        current_first = parse_iso_date(self.settings.get("first_pay_date"))
        if current_first is None or current_first < nxt:
            self.settings["first_pay_date"] = nxt.isoformat()
        archived = {day.isoformat() for day in archived_pay_dates(self.archives)}
        for key in list(self.extras_by_period):
            if str(key)[:10] in archived:
                self.extras_by_period.pop(key, None)

    def _build_charts_tab(self, parent):
        frm = AppFrame(parent, padding=12)
        frm.pack(fill="both", expand=True)
        self._chart_lock = False
        self._debtor_box_choice = None
        self._type_box_choice = None
        self._chart_filter_choice = None
        self._chart_selected_ids = set()
        self._chart_selection_custom = False
        self._chart_model = None

        controls = AppFrame(frm)
        controls.pack(fill="x", pady=(0, 6))
        AppLabel(controls, text="Category", style="Header.TLabel").pack(side="left", padx=(0, 8))
        self.chart_filter_var = tk.StringVar(value="Debt Type")
        self.chart_filter_combo = AppCombo(
            controls, textvariable=self.chart_filter_var, state="readonly", width=16,
            values=("Debt Type", "Debtor"),
        )
        self.chart_filter_combo.pack(side="left")
        self.chart_filter_combo.bind("<<ComboboxSelected>>", self._on_chart_filter)

        self.chart_legend_frame = AppFrame(frm)
        key_bar = AppFrame(self.chart_legend_frame)
        key_bar.pack(fill="x")
        AppLabel(key_bar, text="Click to show or hide", style="OverviewCaption.TLabel").pack(side="left", pady=(2, 4))
        AppButton(key_bar, text="Deselect all", style="ChartKey.TButton", command=self._clear_chart_debts).pack(side="right")
        AppButton(key_bar, text="Select all", style="ChartKey.TButton", command=self._select_all_chart_debts).pack(side="right", padx=(0, 6))
        self._chart_scrollable = False
        self._chart_legend_shown = False
        self.chart_legend_canvas = tk.Canvas(
            self.chart_legend_frame, height=max(32, int(round(28 * self._ui_scale()))),
            bg=COLORS["bg"], highlightthickness=0, cursor="hand2",
        )
        self.chart_legend_canvas._role = "legend"
        self.chart_legend_canvas.pack(fill="x", pady=(2, 0))
        self.chart_legend_canvas.bind("<Button-1>", self._on_chart_legend_click)
        self._chart_frm = frm

        chart_row = AppFrame(frm)
        self._chart_row = chart_row
        chart_row.pack(fill="both", expand=True)
        chart_row.grid_columnconfigure(1, weight=1)
        chart_row.grid_rowconfigure(0, weight=1)
        self.chart_y_canvas = tk.Canvas(
            chart_row, width=int(round(100 * self._ui_scale())), bg=COLORS["chart_bg"], highlightthickness=0,
        )
        self.chart_y_canvas._role = "chart"
        self.chart_y_canvas.grid(row=0, column=0, sticky="ns")
        self.payoff_canvas = tk.Canvas(
            chart_row, bg=COLORS["chart_bg"], highlightthickness=1, highlightbackground=COLORS["chart_frame"],
        )
        self.payoff_canvas._role = "chart"
        self.payoff_canvas.grid(row=0, column=1, sticky="nsew")
        self.payoff_canvas.bind("<Configure>", self._on_chart_resize)

    def _build_interest_tab(self, parent):
        frm = AppFrame(parent, padding=12)
        self.interest_frm = frm
        frm.pack(fill="both", expand=True)
        cards = AppFrame(frm)
        cards.pack(fill="x", anchor="n")
        for column in range(3):
            cards.grid_columnconfigure(column, weight=1, uniform="interest_cards")

        def _box(column, title):
            frame = AppSection(cards, text=title, style="Overview.TLabelframe", padding=10)
            frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0 if column == 2 else 6))
            return frame

        all_box = _box(0, "All Debt")
        AppLabel(all_box, text="Every debt with a balance", style="OverviewCaption.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        debtor_box = _box(1, "By Debtor")
        self.chart_debtor_var = tk.StringVar()
        self.chart_debtor_combo = AppCombo(debtor_box, textvariable=self.chart_debtor_var, state="readonly", width=24)
        self.chart_debtor_combo._combo_fill = "raised"
        self._paint_combo(self.chart_debtor_combo)
        self.chart_debtor_combo.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.chart_debtor_combo.bind("<<ComboboxSelected>>", self._on_chart_debtor_box)
        type_box = _box(2, "By Debt Type")
        self.chart_type_var = tk.StringVar()
        self.chart_type_combo = AppCombo(type_box, textvariable=self.chart_type_var, state="readonly", width=24)
        self.chart_type_combo._combo_fill = "raised"
        self._paint_combo(self.chart_type_combo)
        self.chart_type_combo.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.chart_type_combo.bind("<<ComboboxSelected>>", self._on_chart_type_box)

        self.chart_all_min = tk.StringVar(value="$0.00")
        self.chart_all_add = tk.StringVar(value="$0.00")
        self.chart_all_save = tk.StringVar(value="$0.00")
        self.chart_all_note = tk.StringVar(value="")
        self.chart_debtor_min = tk.StringVar(value="$0.00")
        self.chart_debtor_add = tk.StringVar(value="$0.00")
        self.chart_debtor_save = tk.StringVar(value="$0.00")
        self.chart_debtor_note = tk.StringVar(value="")
        self.chart_type_min = tk.StringVar(value="$0.00")
        self.chart_type_add = tk.StringVar(value="$0.00")
        self.chart_type_save = tk.StringVar(value="$0.00")
        self.chart_type_note = tk.StringVar(value="")
        self._fill_interest_box(all_box, self.chart_all_min, self.chart_all_add, self.chart_all_save, self.chart_all_note)
        self._fill_interest_box(debtor_box, self.chart_debtor_min, self.chart_debtor_add, self.chart_debtor_save, self.chart_debtor_note)
        self._fill_interest_box(type_box, self.chart_type_min, self.chart_type_add, self.chart_type_save, self.chart_type_note)
        for box in (all_box, debtor_box, type_box):
            box.grid_columnconfigure(0, weight=1)

    def _fill_interest_box(self, frame, minimum_var, additional_var, savings_var, note_var):
        AppLabel(frame, text="Interest with minimum payments only", style="OverviewCaption.TLabel").grid(row=1, column=0, sticky="w")
        AppLabel(frame, textvariable=minimum_var, style="OverviewHero.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 8))
        AppLabel(frame, text="Interest with additional payments", style="OverviewCaption.TLabel").grid(row=3, column=0, sticky="w")
        AppLabel(frame, textvariable=additional_var, style="OverviewHero.TLabel").grid(row=4, column=0, sticky="w", pady=(0, 8))
        AppLabel(frame, text="Total savings", style="OverviewCaption.TLabel").grid(row=5, column=0, sticky="w")
        AppLabel(frame, textvariable=savings_var, style="OverviewHero.TLabel").grid(row=6, column=0, sticky="w")
        AppLabel(frame, textvariable=note_var, style="OverviewCaption.TLabel", wraplength=280).grid(row=7, column=0, sticky="w", pady=(6, 0))

    def _on_chart_debtor_box(self, _event=None):
        if self._chart_lock:
            return
        self._debtor_box_choice = self.chart_debtor_var.get()
        self.refresh_charts()

    def _on_chart_type_box(self, _event=None):
        if self._chart_lock:
            return
        self._type_box_choice = self.chart_type_var.get()
        self.refresh_charts()

    def _on_chart_filter(self, _event=None):
        if self._chart_lock:
            return
        self._chart_filter_choice = self.chart_filter_var.get()
        self._chart_selection_custom = False
        self.refresh_charts()

    def _on_chart_legend_click(self, event):
        legend = self.chart_legend_canvas
        x = legend.canvasx(event.x)
        y = legend.canvasy(event.y)
        debt_id = None
        for item_id in legend.find_overlapping(x, y, x, y):
            for tag in legend.gettags(item_id):
                if tag.startswith("debt-"):
                    debt_id = tag[5:]
                    break
            if debt_id is not None:
                break
        if debt_id is None:
            return
        series = (getattr(self, "_chart_model", None) or {}).get("series") or []
        actual = None
        for item in series:
            if str(item.get("id")) == debt_id:
                actual = item["id"]
                break
        if actual is None:
            return
        self._toggle_chart_debt(actual)

    def _toggle_chart_debt(self, debt_id):
        selected = set(self._chart_selected_ids)
        if debt_id in selected:
            selected.remove(debt_id)
        else:
            selected.add(debt_id)
        self._chart_selected_ids = selected
        self._chart_selection_custom = True
        self._draw_payoff_chart(reset_scroll=True)

    def _select_all_chart_debts(self):
        series = (getattr(self, "_chart_model", None) or {}).get("series") or []
        self._chart_selected_ids = {item["id"] for item in series}
        self._chart_selection_custom = True
        self._draw_payoff_chart(reset_scroll=True)

    def _clear_chart_debts(self):
        self._chart_selected_ids = set()
        self._chart_selection_custom = True
        self._draw_payoff_chart(reset_scroll=True)

    def _on_chart_resize(self, event):
        if event.widget is not self.payoff_canvas or getattr(self, "_chart_drawing", False):
            return
        self._draw_payoff_chart()

    def _set_chart_scroll(self, _scroll, _content_width, height, _reset_scroll):
        """Keep the plot inside the canvas. The chart does not scroll."""
        canvas = self.payoff_canvas
        width = max(canvas.winfo_width(), 1)
        canvas.configure(scrollregion=(0, 0, width, max(height, 1)))
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)
        self._chart_scrollable = False

    def _set_chart_legend(self, show, _width=0):
        frame = self.chart_legend_frame
        if show:
            if not self._chart_legend_shown:
                frame.pack(fill="x", pady=(2, 8), before=self._chart_row)
                self._chart_legend_shown = True
        elif self._chart_legend_shown:
            frame.pack_forget()
            self._chart_legend_shown = False

    def _money_text(self, value):
        return f"${value:,.2f}"

    def _interest_note(self, result, debt_ids):
        if any(result["dates"].get(did) == "Never" for did in debt_ids):
            return "Minimum payments do not pay this off within the projection."
        return ""

    def refresh_charts(self):
        if not hasattr(self, "payoff_canvas"):
            return
        today = datetime.date.today()
        snow = simulate_payoff(self.debts, self.settings, today, self.extras_by_period, minimum_only=False)
        minimum = simulate_payoff(self.debts, self.settings, today, self.extras_by_period, minimum_only=True)
        active = _active_debts_in_order(self.debts)
        debtor_names = [debt.get("short_name") or f"Debt {debt.get('id')}" for debt in active]
        type_dates, _overall, by_id = debt_free_dates_by_type(
            self.debts, self.settings, today, self.extras_by_period,
        )
        type_names = []
        for name in OVERVIEW_TYPES:
            members = [debt for debt in active if _debt_type_name(debt) == name]
            if not members:
                continue
            type_names.append((type_dates.get(name) or "Never", name, members))
        type_names.sort(key=lambda item: (item[0] in ("Never", "✓"), item[0]))
        ordered_types = [name for _date, name, _members in type_names]

        def _keep(choice, options, fallback):
            if choice in options:
                return choice
            return fallback

        debtor_choice = _keep(self._debtor_box_choice, debtor_names, debtor_names[0] if debtor_names else "")
        type_choice = _keep(self._type_box_choice, ordered_types, ordered_types[0] if ordered_types else "")
        filter_options = ("Debt Type", "Debtor")
        filter_choice = _keep(self._chart_filter_choice, filter_options, "Debt Type")

        def _ids_for(namesake, group):
            if group == "debtor":
                return [debt["id"] for debt in active if (debt.get("short_name") or f"Debt {debt.get('id')}") == namesake]
            return [debt["id"] for debt in active if _debt_type_name(debt) == namesake]

        def _apply(ids, minimum_var, additional_var, savings_var, note_var):
            minimum_interest = interest_totals(minimum, ids)
            additional_interest = interest_totals(snow, ids)
            minimum_var.set(self._money_text(minimum_interest))
            additional_var.set(self._money_text(additional_interest))
            savings_var.set(self._money_text(minimum_interest - additional_interest))
            note_var.set(self._interest_note(minimum, ids))

        self._chart_lock = True
        try:
            self._debtor_box_choice = debtor_choice
            self._type_box_choice = type_choice
            self._chart_filter_choice = filter_choice
            self.chart_debtor_combo["values"] = debtor_names
            self.chart_type_combo["values"] = ordered_types
            self.chart_debtor_var.set(debtor_choice)
            self.chart_type_var.set(type_choice)
            self.chart_filter_combo["values"] = filter_options
            self.chart_filter_var.set(filter_choice)
            _apply([debt["id"] for debt in active], self.chart_all_min, self.chart_all_add, self.chart_all_save, self.chart_all_note)
            _apply(_ids_for(debtor_choice, "debtor"), self.chart_debtor_min, self.chart_debtor_add, self.chart_debtor_save, self.chart_debtor_note)
            _apply(_ids_for(type_choice, "type"), self.chart_type_min, self.chart_type_add, self.chart_type_save, self.chart_type_note)
            if filter_choice == "Debtor":
                series = chart_debt_lines(snow["points"], self.debts)
            else:
                series = chart_type_lines(snow["points"], self.debts)
            available = {item["id"] for item in series}
            if self._chart_selection_custom:
                kept = {did for did in self._chart_selected_ids if did in available}
                if kept:
                    self._chart_selected_ids = kept
                else:
                    self._chart_selection_custom = False
            if not self._chart_selection_custom:
                smallest = smallest_chart_debt(series)
                self._chart_selected_ids = {smallest["id"]} if smallest else set()
            self._chart_model = {
                "start": snow["start"],
                "frequency": snow["frequency"],
                "series": series,
            }
        finally:
            self._chart_lock = False
        self._draw_payoff_chart(reset_scroll=True)

    def _draw_chart_key(self, series, selected, chart_width):
        """Lay the key in rows above the chart."""
        legend = self.chart_legend_canvas
        legend.delete("all")
        scale = self._ui_scale()
        row_h = max(32, int(round(26 * scale)))
        gap = max(16, int(round(14 * scale)))
        usable = max(legend.winfo_width(), int(chart_width), 240)
        x = 4
        row = 0
        for item in series:
            turned_on = item["id"] in selected
            font = self._tk_font(12, "bold" if turned_on else "normal")
            text_w = tkfont.Font(root=self.root, font=font).measure(str(item["name"]))
            item_w = 24 + text_w
            if x > 4 and x + item_w > usable - 4:
                x = 4
                row += 1
            top = 2 + row * row_h
            center = top + row_h / 2
            tag = f"debt-{item['id']}"
            legend.create_rectangle(x, top, x + item_w, top + row_h - 2, fill=COLORS["bg"], outline="", tags=(tag,))
            legend.create_oval(
                x + 2, center - 7, x + 16, center + 7,
                fill=chart_display_color(item["color"]) if turned_on else COLORS["bg"],
                outline=chart_display_color(item["color"]) if turned_on else COLORS["legend_ring"],
                width=2, tags=(tag,),
            )
            legend.create_text(
                x + 22, center, text=item["name"], anchor="w", tags=(tag,),
                fill=COLORS["text"] if turned_on else COLORS["legend_off"],
                font=font,
            )
            x += item_w + gap
        height = max(row_h, (row + 1) * row_h + 4)
        if abs(int(float(legend.cget("height"))) - height) > 2:
            legend.configure(height=height)

    def _draw_payoff_chart(self, reset_scroll=False):
        canvas = getattr(self, "payoff_canvas", None)
        y_canvas = getattr(self, "chart_y_canvas", None)
        legend = getattr(self, "chart_legend_canvas", None)
        if canvas is None or y_canvas is None or legend is None:
            return
        width = canvas.winfo_width()
        height = canvas.winfo_height()
        if width < 40 or height < 40:
            return
        self._chart_drawing = True
        try:
            canvas.delete("all")
            y_canvas.delete("all")
            legend.delete("all")
            model = getattr(self, "_chart_model", None) or {}
            series = model.get("series") or []
            if not series:
                self._set_chart_legend(False, 1)
                self._set_chart_scroll(False, width, height, True)
                canvas.create_text(
                    width / 2, height / 2, text="No debts with a balance to chart.",
                    fill=COLORS["text"], font=self._tk_font(13),
                )
                return
            selected = set(self._chart_selected_ids)
            visible = [item for item in series if item["id"] in selected]
            self._set_chart_legend(True)
            self._draw_chart_key(series, selected, width)
            start = model.get("start")
            if start is None:
                return
            if not visible:
                self._set_chart_scroll(False, width, height, True)
                canvas.create_text(
                    width / 2, height / 2, text="Click a debt in the key to show it.",
                    fill=COLORS["text"], font=self._tk_font(13),
                )
                return
            scale = self._ui_scale()
            axis_font = self._data_font(12)
            measure = tkfont.Font(root=self.root, font=axis_font)
            end, bulk_days = chart_time_scale(start, visible)
            detailed = bulk_days < 365
            sample = "00/00/00" if detailed else "September '26"
            label_w = max(measure.measure(sample), 48)
            label_h = max(measure.metrics("linespace"), 14)
            angle = -40
            rad = math.radians(abs(angle))
            pitch = int(math.ceil(label_w * math.cos(rad))) + max(14, int(round(12 * scale)))
            drop = int(math.ceil(label_w * math.sin(rad) + label_h * math.cos(rad))) + max(12, int(round(8 * scale)))
            pad_left = int(math.ceil(label_h * math.sin(rad))) + max(10, int(round(6 * scale)))
            pad_right = max(pitch, 24)
            plot_width = max(width - pad_left - pad_right, 40)
            plot_left = pad_left
            plot_right = plot_left + plot_width
            _used, x_ticks = chart_axis_ticks(
                start, end, model.get("frequency") or "biweekly", plot_width,
                detailed=detailed, min_px=pitch,
            )
            plot_top = 16
            plot_bottom = max(plot_top + 40, height - drop)
            plot_height = plot_bottom - plot_top
            self._set_chart_scroll(False, width, height, reset_scroll)
            y_max = 0.0
            for item in visible:
                for _day, amount in item["points"]:
                    if amount > y_max:
                        y_max = amount
            y_top, y_ticks = chart_value_ticks(y_max, plot_height, int(label_h * 1.8))
            span_days = max((end - start).days, 1)

            def x_of(day):
                if day < start:
                    day = start
                if day > end:
                    day = end
                return plot_left + (day - start).days / span_days * plot_width

            def y_of(amount):
                return plot_bottom - (max(0.0, amount) / y_top) * plot_height

            grid = COLORS["chart_grid"]
            wash = COLORS["chart_wash"]
            axis_gap = max(8, int(round(10 * scale)))

            def money_label(value):
                if abs(value) >= 100 or abs(value - round(value)) < 0.05:
                    return f"${value:,.0f}"
                return f"${value:,.2f}"

            y_labels = [money_label(value) for value in y_ticks]
            y_need = max(int(round(100 * scale)), max(measure.measure(text) for text in y_labels) + int(round(18 * scale)))
            if abs(int(float(y_canvas.cget("width"))) - y_need) > 4:
                y_canvas.configure(width=y_need)
            axis_w = max(y_canvas.winfo_width(), y_need, 1)
            canvas.create_rectangle(plot_left, plot_top, plot_right, plot_bottom, outline="", fill=COLORS["chart_plot"])
            y_canvas.create_line(axis_w - 1, plot_top, axis_w - 1, plot_bottom, fill=COLORS["chart_axis"])
            for value, label in zip(y_ticks, y_labels):
                y = y_of(value)
                canvas.create_line(plot_left, y, plot_right, y, fill=grid)
                y_canvas.create_text(axis_w - 10, y, text=label, anchor="e", fill=COLORS["chart_label"], font=axis_font)
            for position, label in x_ticks:
                x = x_of(position)
                canvas.create_line(x, plot_top, x, plot_bottom, fill=grid)
            drawn = []
            for item in visible:
                color = chart_display_color(item["color"])
                coords = []
                for day, amount in item["points"]:
                    if day < start or day > end:
                        continue
                    coords.extend((x_of(day), y_of(amount)))
                if len(coords) >= 4:
                    drawn.append((color, coords))
            if len(drawn) == 1:
                color, coords = drawn[0]
                canvas.create_polygon(
                    *coords, coords[-2], plot_bottom, coords[0], plot_bottom,
                    fill=chart_mix_toward(color, wash, 0.88), outline="",
                )
            for color, coords in drawn:
                if len(coords) >= 4:
                    canvas.create_line(*coords, fill=wash, width=5, capstyle="round", joinstyle="round")
                    canvas.create_line(*coords, fill=color, width=2, capstyle="round", joinstyle="round")
                canvas.create_oval(coords[0] - 3, coords[1] - 3, coords[0] + 3, coords[1] + 3, fill=color, outline=wash, width=1)
                canvas.create_oval(coords[-2] - 4.5, coords[-1] - 4.5, coords[-2] + 4.5, coords[-1] + 4.5, fill=wash, outline=color, width=2)
            canvas.create_rectangle(plot_left, plot_top, plot_right, plot_bottom, outline=COLORS["chart_frame"])
            canvas.create_line(plot_left, plot_bottom, plot_right, plot_bottom, fill=COLORS["chart_axis"])
            canvas.create_line(plot_left, plot_top, plot_left, plot_bottom, fill=COLORS["chart_axis"])
            for position, label in x_ticks:
                x = x_of(position)
                canvas.create_line(x, plot_bottom, x, plot_bottom + max(4, axis_gap // 2), fill=COLORS["chart_axis"])
                canvas.create_text(
                    x, plot_bottom + axis_gap, text=label, angle=angle, anchor="nw",
                    fill=COLORS["chart_label"], font=axis_font,
                )
        finally:
            self._chart_drawing = False

    def recalculate(self):
        if not self.settings.get("balance_as_of"):
            self.settings["balance_as_of"] = self.settings.get("first_pay_date") or datetime.date.today().isoformat()
        self._migrate_extras_to_dates()
        self._clip_live_schedule_after_archives()
        start = schedule_start_date(self.settings, self.archives).isoformat()
        if self.settings.get("first_pay_date") != start:
            self.settings["first_pay_date"] = start
            if hasattr(self, "date_var"):
                self._sync_settings_form()
        try:
            self.schedule = calculate_schedule(self.debts, self.settings, self.extras_by_period, self.archives)
        except Exception:
            self.schedule = []
            self.status_var.set("Could not calculate the schedule. Check dates, amounts, and debt entries.")
            self._refresh_snowball()
            self._mark_views_dirty()
            self._refresh_visible()
            return
        self._refresh_snowball()
        self._mark_views_dirty()
        self._refresh_visible()
        self.status_var.set(f"Schedule calculated — {len(self.schedule)} periods")

    def _refresh_snowball(self):
        if not hasattr(self, "snowball_var"):
            return
        self.snowball_var.set(f"${additional_per_pay(self.debts, self.settings):,.2f}")

    def _hide_paid_off_from_schedule(self):
        """Drop paid-off debts from the pay schedule. A later balance brings one back."""
        retire_paid_off_debts(self.debts)
        assign_payoff_ranks(self.debts)

    def _sync_paid_off_state(self, debt):
        if _is_active_debt(debt):
            self._reopen_paid_off_debt(debt)
            return
        debt.pop("payoff_rank", None)
        debt["schedule_hidden"] = True
        assign_payoff_ranks(self.debts)

    def _reopen_paid_off_debt(self, debt):
        """A paid-off debt given a balance again starts on the current pay period, last in rank."""
        if debt.get("schedule_hidden") and _is_active_debt(debt):
            debt.pop("schedule_hidden", None)
            debt.pop("payoff_rank", None)
            debt["schedule_start"] = current_pay_date(self.settings).isoformat()
            assign_payoff_ranks(self.debts)

    def _live_archive_names(self):
        names = {}
        order = []
        for debt in self._ordered_debts(include_hidden=True):
            try:
                did = int(debt["id"])
            except (TypeError, ValueError):
                did = debt.get("id")
            order.append(did)
            names[did] = debt.get("short_name") or f"Debt {did}"
        return order, names

    def _columns_present_in_period(self, period):
        present = []
        seen = set()
        for source_name in ("balances", "extras", "closings"):
            source = period.get(source_name) or {}
            if not isinstance(source, dict):
                continue
            for key in source:
                try:
                    did = int(key)
                except (TypeError, ValueError):
                    did = key
                if did in seen:
                    continue
                seen.add(did)
                present.append(did)
        return present

    def _freeze_archive_columns(self):
        """Remember the debts in each archived period. Later deletes and renames do not change it."""
        live_order, live_names = self._live_archive_names()
        for periods in self.archives.values():
            for period in periods:
                if "columns" in period:
                    continue
                present = self._columns_present_in_period(period)
                ordered = [did for did in live_order if did in present]
                ordered.extend(did for did in present if did not in ordered)
                period["columns"] = [
                    {"id": did, "short_name": live_names.get(did, f"Debt {did}")}
                    for did in ordered
                ]

    def _columns_for_archived_periods(self, periods):
        self._freeze_archive_columns()
        columns = []
        seen = set()
        for period in periods:
            for column in period.get("columns") or []:
                did = column.get("id")
                if did in seen:
                    continue
                seen.add(did)
                columns.append({"id": did, "short_name": column.get("short_name") or f"Debt {did}"})
        return columns

    def _ordered_debts(self, include_hidden=False):
        debts = self.debts
        if not include_hidden:
            debts = [debt for debt in debts if not debt.get("schedule_hidden")]
        active_debts = [d for d in debts if _is_active_debt(d)]
        paid_debts = [d for d in debts if not _is_active_debt(d)]
        return active_debts + paid_debts

    def _schedule_column_debts(self, include_hidden=False):
        """Pay Schedule columns follow Debts-tab Payoff Rank. Paid-off stay last."""
        debts = self._ordered_debts(include_hidden=include_hidden)
        unpaid = [debt for debt in debts if _is_active_debt(debt)]
        paid = [debt for debt in debts if not _is_active_debt(debt)]
        unpaid.sort(key=_payoff_rank_key)
        return unpaid + paid

    def _make_sheet_header(self, parent, text, anchor, prominent=True):
        scale = self._ui_scale()
        gridline = bool(COLORS.get("gridline_on"))
        justify = {"e": "right", "center": "center"}.get(anchor, "left")
        font = self._data_font(10, "bold")
        padx = max(8, int(round(8 * scale)))
        pady = max(6, int(round(6 * scale)))
        if prominent and COLORS["header_bg"].lower() == COLORS["surface"].lower():
            # Dark header matches the card and the total row. Lift it and mark it with the tab accent.
            bg = COLORS["button_secondary"]
            line = COLORS["tab_line"]
            wrap = tk.Frame(parent, bg=line, bd=0, highlightthickness=0)
            tk.Label(
                wrap, text=text, bg=bg, fg=COLORS["header_fg"], font=font,
                padx=padx, pady=pady, borderwidth=0, relief="flat",
                anchor=anchor, justify=justify,
            ).pack(side="top", fill="both", expand=True)
            tk.Frame(wrap, bg=line, height=max(3, int(round(2 * scale))), bd=0, highlightthickness=0).pack(side="bottom", fill="x")
            return wrap
        return tk.Label(
            parent, text=text, bg=COLORS["header_bg"], fg=COLORS["header_fg"],
            font=font,
            padx=padx, pady=pady,
            borderwidth=0 if gridline else 1,
            relief="flat" if gridline else "solid",
            highlightthickness=1 if gridline else 0,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border"],
            anchor=anchor,
            justify=justify,
        )

    def _sheet_cell_edge(self, bg, outline=None):
        """Keep a 1px edge on every cell so selecting a row does not change its size."""
        if outline:
            return outline
        if COLORS.get("gridline_on") and COLORS.get("gridline"):
            return COLORS["gridline"]
        return bg

    def _make_sheet_cell(self, parent, text, bg, fg, anchor, font, outline=None):
        scale = self._ui_scale()
        edge = self._sheet_cell_edge(bg, outline)
        return tk.Label(
            parent, text=text, bg=bg, fg=fg, font=font,
            padx=max(8, int(round(8 * scale))), pady=max(5, int(round(5 * scale))),
            borderwidth=0, relief="flat",
            highlightthickness=1,
            highlightbackground=edge,
            highlightcolor=edge,
            anchor=anchor,
        )

    def _apply_sheet_cell(self, label, bg, fg, font, outline=None):
        edge = self._sheet_cell_edge(bg, outline)
        label.configure(
            bg=bg, fg=fg, font=font,
            highlightthickness=1,
            highlightbackground=edge,
            highlightcolor=edge,
        )

    def _sheet_row_style(self, index, kind, selected):
        body = self._data_font(10)
        if kind == "total":
            return COLORS["total_bg"], COLORS["total_fg"], self._data_font(10, "bold"), None
        if kind == "paid":
            if selected:
                outline = COLORS.get("current_outline") or COLORS.get("header_bg")
                return COLORS["paid_selected_bg"], COLORS["paid_selected_text"], body, outline
            return COLORS["green"], COLORS["green_text"], body, None
        if selected:
            return (
                COLORS["current_row"], COLORS["text"], body,
                COLORS.get("current_outline") or None,
            )
        if index % 2 == 0:
            return COLORS["alt_row"], COLORS["text"], self._data_font(10), None
        return COLORS["schedule_bg"], COLORS["text"], self._data_font(10), None

    def _sheet_column_floor(self):
        """Same minimum column width the pay schedule uses, so headings stay in their own cells."""
        return max(120, int(round(72 * self._ui_scale())))

    def _pack_sheet(self, parent):
        wrap = AppFrame(parent)
        wrap.pack(fill="both", expand=True)
        sheet = SheetView(wrap, self)
        vsb = AppScroll(wrap, orient="vertical", command=sheet.yview)
        hsb = AppScroll(wrap, orient="horizontal", command=sheet.xview)
        sheet.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        sheet._hsb = hsb
        sheet._vsb = vsb
        sheet.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        return sheet

    def _render_period_grid(self, inner, periods, highlight_current=False, editable=False, include_hidden=False, debts=None, show_year=True, total_debt_last=False, frozen_inner=None, row_surface=None):
        if row_surface is None:
            row_surface = COLORS["schedule_bg"]
        body_canvas = inner.master
        frozen_canvas = frozen_inner.master if frozen_inner is not None else None
        if not periods:
            if frozen_canvas is not None:
                self._paint_period_canvas(frozen_canvas, [], [])
            self._paint_period_canvas(body_canvas, [], [], placeholder="No pay periods in this year.")
            return

        ordered_debts = self._ordered_debts(include_hidden=include_hidden) if debts is None else debts
        columns = [
            {"key": "period", "title": "Period", "frozen": bool(frozen_inner), "header_anchor": "center", "anchor": "center"},
            {"key": "date", "title": "Pay Date", "frozen": bool(frozen_inner), "header_anchor": "center", "anchor": "center"},
        ]
        if show_year:
            columns.append({"key": "year", "title": "Year", "frozen": False, "header_anchor": "center", "anchor": "center"})
        columns.append({"key": "extra", "title": "Total Extra", "frozen": False, "header_anchor": "e", "anchor": "e"})
        if not total_debt_last:
            columns.append({"key": "total", "title": "Total Debt", "frozen": False, "header_anchor": "e", "anchor": "e"})
        for debt in ordered_debts:
            columns.append({"key": "debt", "title": debt["short_name"], "debt": debt, "frozen": False, "header_anchor": "e", "anchor": "e"})
        if total_debt_last:
            columns.append({"key": "total", "title": "Total Debt", "frozen": False, "header_anchor": "e", "anchor": "e"})

        frozen_cols = [col for col in columns if col.get("frozen")]
        body_cols = [col for col in columns if not col.get("frozen")]
        metrics = self._period_metrics()
        body_font = metrics["body"]
        paid_font = metrics["paid"]
        body_spec = metrics["body_spec"]
        paid_spec = metrics["paid_spec"]

        active_idx = None
        if highlight_current:
            active_pay = current_pay_date(self.settings)
            for per in periods:
                if parse_flexible_date(per.get("date")) == active_pay:
                    active_idx = per.get("index")
                    break

        def money_cell(value):
            amount = parse_amount(value, allow_blank=True)
            if amount is None:
                return ""
            return f"${amount:,.2f}"

        def plain_value(spec, period):
            key = spec["key"]
            if key == "period":
                value = period.get("period")
                return "" if value is None else str(value)
            if key == "date":
                return format_display_date(period.get("date") or "")
            if key == "year":
                value = period.get("year")
                return "" if value is None else str(value)
            if key == "extra":
                return money_cell(period.get("total_extra"))
            if key == "total":
                return money_cell(period.get("total_debt"))
            return ""

        def open_extras(period_index, debt_id=None):
            if not editable or period_index is None:
                return None
            if debt_id is None:
                return lambda idx=period_index: self.edit_extras_for_period(idx)
            return lambda idx=period_index, did=debt_id: self.edit_extras_for_period(idx, focus_debt_id=did)

        frozen_rows = []
        body_rows = []
        for row_number, period in enumerate(periods, start=1):
            is_current = highlight_current and period.get("index") == active_idx
            row_bg = COLORS["current_row"] if is_current else (row_surface if row_number % 2 == 0 else COLORS["alt_row"])
            outline = (COLORS.get("current_outline") or None) if is_current else None
            period_index = period.get("index")

            def pack(text, bg, fg, anchor, font_obj, font_spec, callback):
                return {
                    "text": text,
                    "bg": bg,
                    "fg": fg,
                    "anchor": anchor,
                    "font_obj": font_obj,
                    "font_spec": font_spec,
                    "callback": callback,
                    "outline": outline,
                }

            if frozen_cols:
                frozen_rows.append([
                    pack(
                        plain_value(spec, period), row_bg, COLORS["text"], spec["anchor"],
                        body_font, body_spec, open_extras(period_index),
                    )
                    for spec in frozen_cols
                ])
            body_row = []
            for spec in body_cols:
                if spec["key"] == "debt":
                    debt = spec["debt"]
                    raw = mapping_get(period.get("balances"), debt["id"], None)
                    amount = None if raw is None else parse_amount(raw, allow_blank=True)
                    if amount is None:
                        text, cell_bg, cell_fg, anchor = "", row_bg, COLORS["text"], "center"
                        font_obj, font_spec = body_font, body_spec
                    elif amount < 0.5:
                        text, cell_bg, cell_fg, anchor = "✓", COLORS["green"], COLORS["green_text"], "e"
                        font_obj, font_spec = paid_font, paid_spec
                    else:
                        text, cell_bg, cell_fg, anchor = f"${amount:,.2f}", row_bg, COLORS["text"], "e"
                        font_obj, font_spec = body_font, body_spec
                    body_row.append(pack(
                        text, cell_bg, cell_fg, anchor, font_obj, font_spec,
                        open_extras(period_index, debt["id"]),
                    ))
                else:
                    cell_fg = COLORS["text"]
                    if spec["key"] == "extra":
                        applied = parse_amount(period.get("total_extra"), allow_blank=True)
                        if applied is not None and applied > 0.004:
                            cell_fg = COLORS["extra_text"]
                    body_row.append(pack(
                        plain_value(spec, period), row_bg, cell_fg, spec["anchor"],
                        body_font, body_spec, open_extras(period_index),
                    ))
            body_rows.append(body_row)

        if frozen_canvas is not None:
            self._paint_period_canvas(frozen_canvas, frozen_cols, frozen_rows)
        self._paint_period_canvas(body_canvas, body_cols, body_rows)

    def _period_metrics(self):
        scale = self._ui_scale()
        body = self._grid_font(10)
        header = self._grid_font(10, "bold")
        paid = self._grid_font(12, "bold")
        pad_x = max(8, int(round(8 * scale)))
        pad_y = max(6, int(round(6 * scale)))
        cell_pad_y = max(5, int(round(5 * scale)))
        prominent = COLORS["header_bg"].lower() == COLORS["surface"].lower()
        bar = max(3, int(round(2 * scale))) if prominent else 0
        header_h = header.metrics("linespace") + pad_y * 2 + bar
        cell_h = max(
            28,
            body.metrics("linespace") + cell_pad_y * 2,
            paid.metrics("linespace") + cell_pad_y * 2,
        )
        return {
            "body": body,
            "header": header,
            "paid": paid,
            "body_spec": self._data_font(10),
            "header_spec": self._data_font(10, "bold"),
            "paid_spec": self._data_font(12, "bold"),
            "pad_x": pad_x,
            "bar": bar,
            "header_h": header_h,
            "cell_h": cell_h,
            "prominent": prominent,
        }

    def _text_anchor_x(self, left, width, pad, anchor):
        if anchor == "e":
            return left + width - pad, "e"
        if anchor == "center":
            return left + width / 2, "center"
        return left + pad, "w"

    def _store_period_size(self, canvas, size):
        if canvas is getattr(self, "sched_frozen_canvas", None):
            self._sched_frozen_size = size
        elif canvas is getattr(self, "sched_canvas", None):
            self._sched_body_size = size
        elif canvas is getattr(self, "archive_canvas", None):
            self._archive_size = size

    def _prepare_period_canvas(self, canvas):
        if not getattr(canvas, "_period_window_hidden", False):
            for item in canvas.find_all():
                if canvas.type(item) == "window":
                    canvas.itemconfigure(item, state="hidden")
            canvas._period_window_hidden = True
        if not getattr(canvas, "_period_click_bound", False):
            canvas.bind("<Double-Button-1>", self._on_period_canvas_double)
            canvas._period_click_bound = True
        if not getattr(canvas, "_period_configure_bound", False):
            canvas.bind("<Configure>", lambda _event, target=canvas: self._redraw_visible_rows(target), add="+")
            canvas._period_configure_bound = True

    def _archive_yview(self, *args):
        canvas = getattr(self, "archive_canvas", None)
        if canvas is None:
            return
        canvas.yview(*args)
        self._redraw_visible_rows(canvas)

    def _redraw_visible_rows(self, canvas):
        model = getattr(canvas, "_period_model", None)
        if not model:
            return
        try:
            if not canvas.winfo_exists():
                return
            canvas.delete("period-row")
            view_top = canvas.canvasy(0)
            view_bottom = view_top + max(canvas.winfo_height(), model["cell_h"])
        except tk.TclError:
            return
        rows = model["rows"]
        widths = model["widths"]
        header_h = model["header_h"]
        cell_h = model["cell_h"]
        total_w = model["total_w"]
        pad_x = model["pad_x"]
        grid = model["grid"]
        if not rows or cell_h <= 0:
            canvas._period_hits = []
            if self._period_header_sticks(canvas):
                self._draw_period_header(canvas, view_top)
            return
        first = max(0, int((view_top - header_h) / cell_h) - 1)
        last = min(len(rows), int((view_bottom - header_h) / cell_h) + 2)
        hits = []
        for index in range(first, last):
            row = rows[index]
            top = header_h + index * cell_h
            row_bg = row[0]["bg"] if row else COLORS["schedule_bg"]
            canvas.create_rectangle(0, top, total_w, top + cell_h, fill=row_bg, outline="", tags="period-row")
            left = 0
            for width, cell in zip(widths, row):
                if cell["bg"] != row_bg:
                    canvas.create_rectangle(
                        left, top, left + width, top + cell_h,
                        fill=cell["bg"], outline="", tags="period-row",
                    )
                text_x, text_anchor = self._text_anchor_x(left, width, pad_x, cell["anchor"])
                canvas.create_text(
                    text_x, top + cell_h / 2, text=cell["text"], anchor=text_anchor,
                    fill=cell["fg"], font=cell["font_spec"], tags="period-row",
                )
                if cell.get("callback"):
                    hits.append((left, top, left + width, top + cell_h, cell["callback"]))
                left += width
            if grid:
                canvas.create_line(0, top + cell_h, total_w, top + cell_h, fill=grid, tags="period-row")
            outline = row[0].get("outline") if row else None
            if outline:
                canvas.create_rectangle(
                    1, top + 1, max(total_w - 1, 2), top + cell_h - 1,
                    outline=outline, width=2, tags="period-row",
                )
        if grid and last > first:
            band_top = header_h + first * cell_h
            band_bottom = header_h + last * cell_h
            left = 0
            canvas.create_line(0, band_top, 0, band_bottom, fill=grid, tags="period-row")
            for width in widths:
                left += width
                canvas.create_line(left, band_top, left, band_bottom, fill=grid, tags="period-row")
        canvas._period_hits = hits
        if self._period_header_sticks(canvas):
            self._draw_period_header(canvas, view_top)

    def _period_header_sticks(self, canvas):
        """Pay Schedule titles stay at the top of the view. Archives keep a scrolling title."""
        return canvas in (
            getattr(self, "sched_canvas", None),
            getattr(self, "sched_frozen_canvas", None),
        )

    def _draw_period_header(self, canvas, top):
        model = getattr(canvas, "_period_model", None)
        if not model or not model.get("headers"):
            return
        canvas.delete("period-header")
        header_h = model["header_h"]
        total_w = model["total_w"]
        pad_x = model["pad_x"]
        bar = model["bar"]
        canvas.create_rectangle(
            0, top, total_w, top + header_h,
            fill=model["header_bg"], outline="", tags=("period", "period-header"),
        )
        left = 0
        text_y = top + (header_h - bar) / 2
        for spec in model["headers"]:
            width = spec["width"]
            text_x, text_anchor = self._text_anchor_x(left, width, pad_x, spec["anchor"])
            canvas.create_text(
                text_x, text_y, text=spec["title"], anchor=text_anchor,
                fill=model["header_fg"], font=model["header_font"],
                tags=("period", "period-header"),
            )
            left += width
        if bar:
            canvas.create_rectangle(
                0, top + header_h - bar, total_w, top + header_h,
                fill=model["bar_color"], outline="", tags=("period", "period-header"),
            )
        canvas.tag_raise("period-header")

    def _on_period_canvas_double(self, event):
        canvas = event.widget
        x = canvas.canvasx(event.x)
        y = canvas.canvasy(event.y)
        if self._period_header_sticks(canvas):
            model = getattr(canvas, "_period_model", None) or {}
            header_h = model.get("header_h") or 0
            if y < canvas.canvasy(0) + header_h:
                return
        for left, top, right, bottom, callback in getattr(canvas, "_period_hits", ()):
            if callback and left <= x < right and top <= y < bottom:
                callback()
                return

    def _paint_period_canvas(self, canvas, columns, rows, placeholder=""):
        self._prepare_period_canvas(canvas)
        canvas.delete("period")
        try:
            canvas.configure(bg=COLORS["schedule_bg"])
        except tk.TclError:
            return
        metrics = self._period_metrics()
        if not columns:
            if placeholder:
                canvas.create_text(
                    12, 16, text=placeholder, anchor="nw", fill=COLORS["text"],
                    font=self._data_font(11), tags="period",
                )
            canvas._period_hits = []
            canvas._period_model = None
            size = (320, 48) if placeholder else (1, 1)
            self._store_period_size(canvas, size)
            canvas.configure(scrollregion=(0, 0, size[0], size[1]))
            return
        pad_x = metrics["pad_x"]
        floor = self._sheet_column_floor()
        header_font = metrics["header"]
        body_font = metrics["body"]
        widths = []
        for spec in columns:
            key = spec.get("key")
            if key == "period":
                sample = "999"
            elif key == "date":
                sample = "12-31-2026"
            elif key == "year":
                sample = "2026"
            else:
                sample = "$999,999.99"
            widths.append(max(
                floor,
                header_font.measure(str(spec["title"])) + pad_x * 2,
                body_font.measure(sample) + pad_x * 2,
            ))
        header_h = metrics["header_h"]
        cell_h = metrics["cell_h"]
        total_w = max(sum(widths), 1)
        total_h = header_h + cell_h * max(len(rows), 0)
        header_bg = COLORS["button_secondary"] if metrics["prominent"] else COLORS["header_bg"]
        grid = COLORS["gridline"] if COLORS.get("gridline_on") else ""
        canvas._period_model = {
            "widths": widths,
            "rows": rows,
            "header_h": header_h,
            "cell_h": cell_h,
            "total_w": total_w,
            "pad_x": pad_x,
            "grid": grid,
            "headers": [
                {"title": spec["title"], "anchor": spec.get("header_anchor", "w"), "width": width}
                for width, spec in zip(widths, columns)
            ],
            "header_bg": header_bg,
            "header_fg": COLORS["header_fg"],
            "header_font": metrics["header_spec"],
            "bar": metrics["bar"],
            "bar_color": COLORS["tab_line"],
        }
        canvas._period_hits = []
        size = (total_w, max(total_h, 1))
        self._store_period_size(canvas, size)
        canvas.configure(scrollregion=(0, 0, size[0], size[1]))
        if not self._period_header_sticks(canvas):
            self._draw_period_header(canvas, 0)
        self._redraw_visible_rows(canvas)
        self.root.after_idle(lambda target=canvas: self._redraw_visible_rows(target))

    def refresh_schedule_tree(self):
        y = 0
        try:
            y = self.sched_canvas.yview()[0]
        except tk.TclError:
            y = 0
        self._render_period_grid(
            self.sched_inner,
            self.schedule,
            highlight_current=True,
            editable=True,
            show_year=False,
            total_debt_last=True,
            frozen_inner=self.sched_frozen_inner,
            debts=self._schedule_column_debts(),
        )
        try:
            self.sched_canvas.yview_moveto(y)
            self.sched_frozen_canvas.yview_moveto(y)
            self._redraw_visible_rows(self.sched_canvas)
            self._redraw_visible_rows(self.sched_frozen_canvas)
        except tk.TclError:
            pass
        fit = getattr(self, "_fit_schedule_panes", None)
        if fit:
            fit()
        dirty = getattr(self, "_views_dirty", None)
        if dirty:
            dirty.discard("Pay Schedule")

    def edit_extras_for_period(self, idx, focus_debt_id=None):
        if self._block_sample_edit():
            return
        if idx >= len(self.schedule):
            return
        period = self.schedule[idx]
        win = ctk.CTkToplevel(self.root)
        self._prepare_dialog(win)
        title_extra = ""
        if focus_debt_id:
            dname = next((d["short_name"] for d in self.debts if d["id"] == focus_debt_id), "")
            title_extra = f" — {dname}"
        win.title(f"Extras for Period {period['period']} ({format_display_date(period['date'])}){title_extra}")
        win.geometry("480x520")
        win.configure(fg_color=COLORS["bg"])

        AppLabel(win, text=f"Assign extra payments for {format_display_date(period['date'])}", style="Header.TLabel").pack(pady=8)
        entries = {}
        entry_widgets = {}
        closing = False

        bottom = AppFrame(win)
        bottom.pack(side="bottom", fill="x", pady=(4, 12))
        AppLabel(bottom, text="A valid extra is applied when you leave the field or press Done.", style="OverviewCaption.TLabel").pack(pady=(0, 8))
        done_btn = AppButton(bottom, text="Done", command=lambda: close_dialog())
        done_btn.pack()
        done_btn.bind("<ButtonPress-1>", lambda _event: remember_focused_entry(), add="+")

        frame = AppFrame(win)
        frame.pack(fill="both", expand=True, padx=12)

        visible_debts = self._schedule_column_debts()
        if not visible_debts:
            AppLabel(frame, text="No debts on the pay schedule.").grid(row=0, column=0, sticky="w", pady=3)

        for i, d in enumerate(visible_debts):
            did = d["id"]
            period_map = mapping_get(self.extras_by_period, period["date"], {})
            if not isinstance(period_map, dict):
                period_map = {}
            current_extra = mapping_get(period_map, did, 0.0) or 0.0
            label_text = d["short_name"]
            if focus_debt_id and did == focus_debt_id:
                label_text = f"★ {label_text}"
            AppLabel(frame, text=label_text).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(current_extra))
            ent = AppEntry(frame, textvariable=var, width=14)
            ent.grid(row=i, column=1, padx=6, pady=3)
            entries[did] = var
            entry_widgets[did] = ent
            if focus_debt_id and did == focus_debt_id:
                ent.focus_set()
                ent.select_range(0, "end")
            ent.bind("<FocusOut>", lambda _event: commit())
            ent.bind("<Return>", lambda _event: commit())

        def remember_focused_entry():
            """Copy the entry that still has focus. Done is pressed before that field loses focus."""
            try:
                focused = win.focus_get()
            except tk.TclError:
                return
            for did, ent in entry_widgets.items():
                if ent == focused:
                    entries[did].set(ent.get())
                    return

        def commit():
            nonlocal closing
            if closing or not win.winfo_exists():
                return True
            remember_focused_entry()
            updated = {}
            for did, var in entries.items():
                text = var.get().strip()
                if not text:
                    continue
                amount = parse_amount(text, minimum=0, maximum=1e12)
                if amount is None:
                    self.status_var.set("Enter a non-negative number for each extra payment")
                    return False
                if amount > 0:
                    updated[did] = amount
            current = mapping_get(self.extras_by_period, period["date"], {})
            if not isinstance(current, dict):
                current = {}
            current = {int(key) if str(key).isdigit() else key: value for key, value in current.items()}
            shown_ids = set(entries)
            for key, value in current.items():
                if key in shown_ids:
                    continue
                amount = parse_amount(value, minimum=0, maximum=1e12)
                if amount and amount > 0:
                    updated[key] = amount
            if current == updated:
                return True
            self.push_undo()
            if updated:
                self.extras_by_period[period["date"]] = updated
            else:
                self.extras_by_period.pop(period["date"], None)
            self.recalculate()
            return True

        def close_dialog():
            nonlocal closing
            if closing or not win.winfo_exists():
                return
            remember_focused_entry()
            if not commit():
                return
            closing = True
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", close_dialog)

    def archive_years(self):
        """Archive pay periods before the one still in effect. Manual only — no auto-archive."""
        if self._block_sample_edit():
            return
        today = datetime.date.today()
        active = current_pay_date(self.settings, today)
        already = archived_pay_dates(self.archives)
        to_archive = []
        remaining = []
        for p in self.schedule:
            try:
                pd = datetime.datetime.strptime(p["date"], "%Y-%m-%d").date()
            except Exception:
                pd = today
            if pd in already:
                continue
            if pd < active:
                to_archive.append(p)
            else:
                remaining.append(p)

        if not to_archive:
            self._clip_live_schedule_after_archives()
            self.recalculate()
            if already:
                messagebox.showinfo(
                    "Archive",
                    "Those completed pay periods are already in Archives. "
                    "They were removed from the live Pay Schedule so they are not archived twice.",
                )
            else:
                messagebox.showinfo("Archive", "No completed pay periods to archive.")
            return

        self.push_undo()
        live_order, live_names = self._live_archive_names()
        for p in to_archive:
            year = p["year"]
            if year not in self.archives:
                self.archives[year] = []
            snapshot = deepcopy(p)
            present = self._columns_present_in_period(snapshot)
            ordered = [did for did in live_order if did in present]
            ordered.extend(did for did in present if did not in ordered)
            snapshot["columns"] = [
                {"id": did, "short_name": live_names.get(did, f"Debt {did}")}
                for did in ordered
            ]
            self.archives[year].append(snapshot)
            self.extras_by_period.pop(p["date"], None)

        if remaining:
            first = remaining[0]
            last = to_archive[-1]
            for debt in self.debts:
                baked = mapping_get(last.get("closings") or last["balances"], debt["id"], None)
                baked_value = parse_amount(baked, minimum=0, allow_blank=True) if baked is not None else None
                if baked_value is not None:
                    debt["current"] = baked_value
                    debt.pop("schedule_start", None)
                    if baked_value < 0.5:
                        debt["schedule_hidden"] = True
                    else:
                        debt.pop("schedule_hidden", None)
            self.settings["balance_as_of"] = first["date"]
            self.settings["first_pay_date"] = first["date"]
            self.settings["earliest_pay_date"] = first["date"]
            self._sync_settings_form()

        self.refresh_debt_tree()
        self.recalculate()
        self.refresh_archives()
        messagebox.showinfo("Archived", f"Archived {len(to_archive)} completed pay period(s).")

    # ----- Archives -----
    def refresh_archives(self):
        selected = getattr(self, "_archive_year", None)
        for child in self.archive_year_bar.winfo_children():
            child.destroy()
        years = sorted(self.archives.keys(), key=lambda year: int(year), reverse=True)
        if not years:
            self._archive_year = None
            AppLabel(self.archive_year_bar, text="No archived years yet.", style="OverviewCaption.TLabel").pack(side="left", padx=2, pady=4)
            if hasattr(self, "archive_inner"):
                self._render_period_grid(self.archive_inner, [])
            return
        if selected is None or not any(int(year) == int(selected) for year in years):
            selected = int(years[0])
            self._archive_year = selected
        for year in years:
            year_number = int(year)
            chosen = year_number == int(selected)
            button = ctk.CTkButton(
                self.archive_year_bar,
                text=str(year_number),
                font=ui_font( 13, "bold" if chosen else "normal"),
                fg_color=COLORS["button"] if chosen else COLORS["button_secondary"],
                hover_color=COLORS["button_hover"] if chosen else COLORS["button_secondary_hover"],
                text_color=COLORS["button_text"] if chosen else COLORS["button_secondary_text"],
                border_width=1,
                border_color=COLORS["button"] if chosen else COLORS["button_secondary_border"],
                corner_radius=6,
                height=30,
                width=72,
                command=lambda picked=year_number: self._select_archive_year(picked),
            )
            button.pack(side="left", padx=(0, 6), pady=2)
        self._show_archive_year(int(selected))

    def _select_archive_year(self, year):
        self._archive_year = int(year)
        self.refresh_archives()

    def _show_archive_year(self, year):
        periods = list(self.archives.get(year, self.archives.get(str(year), [])))
        periods.sort(key=lambda period: period.get("date", ""))
        self._render_period_grid(
            self.archive_inner,
            periods,
            highlight_current=False,
            editable=False,
            debts=self._columns_for_archived_periods(periods),
        )

    # ----- File I/O -----
    def new_blank(self):
        if messagebox.askyesno("New Blank", "Clear all debts and start blank?"):
            self._start_blank_plan()

    def _start_blank_plan(self):
        mode = self.settings.get("appearance", "dark")
        self.debts = []
        self.settings = deepcopy(DEFAULT_SETTINGS)
        self.settings["appearance"] = mode
        today = datetime.date.today().isoformat()
        self.settings["first_pay_date"] = today
        self.settings["balance_as_of"] = today
        self.settings.pop("earliest_pay_date", None)
        self.extras_by_period = {}
        self.archives = {}
        self.current_file = None
        self._forget_remembered_file()
        self._leave_sample_view()
        self._set_title()
        self._sync_settings_form()
        self.refresh_debt_tree()
        self.recalculate()
        self.refresh_archives()
        self._mark_saved_state()
        self.status_var.set("Started blank")

    def _startup_prompts(self):
        if self.current_file or os.path.isfile(self._own_marker_path()):
            self._schedule_update_check()
            return
        if os.path.isfile(self._sample_marker_path()):
            self._open_sample_plan()
            self._schedule_update_check()
            return
        self._prompt_first_plan()
        self._schedule_update_check()

    def _schedule_update_check(self):
        try:
            if self.root.winfo_exists():
                self.root.after(400, lambda: self.check_for_updates(interactive=False))
        except tk.TclError:
            return

    def check_for_updates(self, interactive=False):
        """Ask GitHub which version is published. Startup stays quiet unless a newer one exists."""
        if getattr(self, "_update_check_running", False):
            if interactive:
                self.status_var.set("Already checking for updates")
            return
        self._update_check_running = True
        self._update_result = None
        if interactive:
            self.status_var.set("Checking GitHub for updates...")
        token = object()
        self._update_token = token

        def work():
            remote = None
            error = None
            try:
                remote = fetch_published_version()
            except Exception:
                error = "Could not reach GitHub to check for updates."
            if getattr(self, "_update_token", None) is token:
                self._update_result = (remote, error, interactive)

        def poll():
            if getattr(self, "_update_token", None) is not token:
                return
            try:
                if not self.root.winfo_exists():
                    self._update_check_running = False
                    return
            except tk.TclError:
                self._update_check_running = False
                return
            result = self._update_result
            if result is None:
                self.root.after(100, poll)
                return
            self._update_result = None
            self._update_check_running = False
            remote, error, asked = result
            self._show_update_result(remote, error, asked)

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, poll)

    def _show_update_result(self, remote, error, interactive):
        if error or not remote:
            if interactive:
                messagebox.showwarning("Check for Updates", error or "Could not read the version on GitHub.", parent=self.root)
                self.status_var.set("Update check failed")
            return
        if version_is_newer(remote, APP_VERSION):
            self.status_var.set(f"Version {remote} is available")
            if sys.platform.startswith("win"):
                prompt = (
                    f"Version {remote} is available. You have version {APP_VERSION}.\n\n"
                    "Download the installer now? Your saved plan stays on this computer."
                )
            else:
                prompt = (
                    f"Version {remote} is available. You have version {APP_VERSION}.\n\n"
                    "Open the install page for your system?"
                )
            if messagebox.askyesno("Update available", prompt, default=messagebox.NO, parent=self.root):
                webbrowser.open(update_page_url(remote))
            return
        if interactive:
            messagebox.showinfo("Check for Updates", f"You have the latest version ({APP_VERSION}).", parent=self.root)
            self.status_var.set(f"Version {APP_VERSION} is up to date")

    def _prompt_first_plan(self):
        choice = self._plan_choice_dialog()
        if choice == "sample":
            self._open_sample_plan()
            return
        if choice == "new":
            self._start_blank_plan()
            return
        try:
            if self.root.winfo_exists():
                self.root.destroy()
        except tk.TclError:
            pass

    def _plan_choice_dialog(self):
        choice = {"value": None}
        win = ctk.CTkToplevel(self.root)
        self._prepare_dialog(win)
        win.title("Welcome")
        win.geometry("560x200")
        win.configure(fg_color=COLORS["bg"])
        win.transient(self.root)
        AppLabel(
            win,
            text="View the included sample plan, or create your own. The sample plan cannot be edited.",
            wraplength=500,
            justify="left",
        ).pack(padx=20, pady=(20, 16), anchor="w")
        buttons = AppFrame(win)
        buttons.pack(pady=8)

        def choose_sample():
            choice["value"] = "sample"
            win.destroy()

        def choose_new():
            choice["value"] = "new"
            win.destroy()

        def choose_close():
            choice["value"] = None
            win.destroy()

        AppButton(buttons, text="View Sample Plan", style="secondary", command=choose_sample).pack(side="left", padx=6)
        AppButton(buttons, text="Create New Plan", command=choose_new).pack(side="left", padx=6)
        win.protocol("WM_DELETE_WINDOW", choose_close)
        win.grab_set()
        win.wait_window()
        return choice["value"]

    def _user_data_dir(self):
        """Folder for the remembered plan path. An installed copy uses AppData."""
        if getattr(sys, "frozen", False):
            root = os.environ.get("APPDATA") or os.path.expanduser("~")
            folder = os.path.join(root, "Debt Payoff Planner")
            os.makedirs(folder, exist_ok=True)
            return folder
        if not sys.platform.startswith("win"):
            # A deb or Flatpak copy lives in a read-only install folder.
            root = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
            folder = os.path.join(root, "debt-payoff-planner")
            os.makedirs(folder, exist_ok=True)
            return folder
        return os.path.dirname(os.path.abspath(__file__))

    def _sample_marker_path(self):
        return os.path.join(self._user_data_dir(), "viewing_sample.txt")

    def _own_marker_path(self):
        return os.path.join(self._user_data_dir(), "own_plan.txt")

    def _bundled_sample_path(self):
        candidates = []
        if getattr(sys, "frozen", False):
            candidates.append(os.path.join(getattr(sys, "_MEIPASS", ""), "sample_plan.json"))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_plan.json"))
        for path in candidates:
            if path and os.path.isfile(path):
                return path
        return None

    def _block_sample_edit(self):
        if not getattr(self, "sample_view", False):
            return False
        self.status_var.set("Sample plan is view only. Create or open your own plan to make changes.")
        return True

    def _set_sample_controls(self, enabled):
        state = "normal" if enabled else "disabled"
        frame = getattr(self, "debt_btn_frame", None)
        if frame is not None:
            for child in frame.winfo_children():
                if isinstance(child, ctk.CTkButton):
                    child.configure(state=state)
        archive = getattr(self, "sched_archive_btn", None)
        if archive is not None:
            archive.configure(state=state)
        for name in (
            "settings_appearance", "settings_first_pay", "settings_balance_as_of",
            "settings_starting", "settings_proj", "settings_salary", "settings_freq",
            "settings_spouse", "settings_spouse_freq", "settings_other", "settings_other_freq",
        ):
            widget = getattr(self, name, None)
            if widget is None:
                continue
            widget_state = state
            if enabled and isinstance(widget, ctk.CTkComboBox):
                widget_state = "readonly"
            try:
                widget.configure(state=widget_state)
            except tk.TclError:
                pass
        menu = getattr(self, "_file_menu", None)
        if menu is not None:
            for label in ("Undo", "Redo"):
                try:
                    menu.entryconfig(label, state=state)
                except tk.TclError:
                    pass

    def _enter_sample_view(self):
        self.sample_view = True
        self.current_file = None
        try:
            with open(self._sample_marker_path(), "w", encoding="utf-8") as handle:
                handle.write("1")
        except OSError:
            pass
        self._set_sample_controls(False)
        self._set_title()

    def _leave_sample_view(self):
        self.sample_view = False
        try:
            os.remove(self._sample_marker_path())
        except OSError:
            pass
        try:
            with open(self._own_marker_path(), "w", encoding="utf-8") as handle:
                handle.write("1")
        except OSError:
            pass
        self._set_sample_controls(True)

    def _open_sample_plan(self):
        path = self._bundled_sample_path()
        if not path:
            messagebox.showerror("Sample plan", "The included sample plan could not be found.")
            return
        try:
            self._load_saved_plan(path, remember=False)
        except Exception as error:
            messagebox.showerror("Sample plan", f"Could not open the sample plan:\n{error}")

    def _recent_file_path(self):
        return os.path.join(self._user_data_dir(), "last_plan.txt")

    def _remember_file(self, path):
        try:
            with open(self._recent_file_path(), "w", encoding="utf-8") as handle:
                handle.write(os.path.abspath(path))
        except OSError:
            pass

    def _forget_remembered_file(self):
        try:
            os.remove(self._recent_file_path())
        except OSError:
            pass

    def _set_title(self):
        if getattr(self, "sample_view", False):
            self.root.title("Debt Payoff Planner — Sample Plan")
        elif self.current_file:
            self.root.title(f"Debt Payoff Planner — {os.path.basename(self.current_file)}")
        else:
            self.root.title("Debt Payoff Planner")

    def _load_saved_plan(self, path, remember=True):
        with open(path, "r") as handle:
            data = json.load(handle)
        if not isinstance(data, dict):
            raise ValueError("Plan file is not a JSON object")
        cleaned_debts = []
        for debt in data.get("debts") or []:
            item = sanitize_debt(debt)
            if item is None:
                continue
            normalize_debt_due(item)
            cleaned_debts.append(item)
        self.debts = cleaned_debts
        assign_payoff_ranks(self.debts)
        self.settings = sanitize_settings(data.get("settings"))
        self._apply_appearance(refresh=False)
        self.extras_by_period = sanitize_extras(data.get("extras_by_period", {}))
        self.archives = normalize_archives(data.get("archives", {}))
        self._freeze_archive_columns()
        self.current_file = path if remember else None
        if remember:
            self._remember_file(path)
            self._leave_sample_view()
        else:
            rebase_sample_calendar(self.settings, self.debts, self.extras_by_period, self.archives)
        self._set_title()
        self._sync_settings_form()
        self._hide_paid_off_from_schedule()
        self.refresh_debt_tree()
        self.recalculate()
        self.refresh_archives()
        self._mark_saved_state()
        if remember:
            self.status_var.set(f"Opened {os.path.basename(path)}")
        else:
            self._enter_sample_view()
            self.status_var.set("Sample plan — view only")

    def _load_remembered_file(self):
        try:
            with open(self._recent_file_path(), "r", encoding="utf-8") as handle:
                path = handle.read().strip()
        except OSError:
            return False
        if not path or not os.path.isfile(path):
            return False
        try:
            self._load_saved_plan(path)
            return True
        except Exception:
            self.status_var.set("Could not reopen the last saved plan")
            return False

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return False
        try:
            self._load_saved_plan(path)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
            return False

    def _persisted_state(self):
        return {
            "debts": self.debts,
            "settings": self.settings,
            "extras_by_period": self.extras_by_period,
            "archives": self.archives,
        }

    def _mark_saved_state(self):
        self._saved_token = json.dumps(self._persisted_state(), sort_keys=True, default=str)

    def _has_unsaved_changes(self):
        try:
            return json.dumps(self._persisted_state(), sort_keys=True, default=str) != self._saved_token
        except TypeError:
            return True

    def confirm_close(self):
        if self._has_unsaved_changes():
            answer = messagebox.askyesnocancel(
                "Unsaved Changes",
                "Save changes to your plan before closing?",
            )
            if answer is None:
                return
            if answer and not self.save_file():
                return
        self.root.destroy()

    def save_file(self):
        if self.current_file:
            if not self._write_json(self.current_file):
                return False
            self.status_var.set(f"Saved {os.path.basename(self.current_file)}")
            return True
        return self.save_file_as()

    def save_file_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not path:
            return False
        previous = self.current_file
        self.current_file = path
        if not self._write_json(path):
            self.current_file = previous
            self._set_title()
            return False
        self._leave_sample_view()
        self._set_title()
        self.status_var.set(f"Saved {os.path.basename(path)}")
        return True

    def _write_json(self, path):
        data = self._persisted_state()
        try:
            with open(path, "w") as f:
                json.dump(data, f, indent=2, allow_nan=False)
        except OSError as e:
            messagebox.showerror("Save failed", f"Could not save file:\n{e}")
            return False
        self._mark_saved_state()
        self._remember_file(path)
        self._set_title()
        return True

    # ----- Excel export -----
    def export_schedule(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing library", "Install openpyxl first:\n  pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Pay Schedule"

        # Headers
        shown_debts = self._schedule_column_debts()
        headers = ["Period", "Date", "Year", "Total Extra", "Total Debt"] + [d["short_name"] for d in shown_debts]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(1, col, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2C3E50")

        for row, p in enumerate(self.schedule, 2):
            ws.cell(row, 1, p["period"])
            ws.cell(row, 2, format_display_date(p["date"]))
            ws.cell(row, 3, p["year"])
            ws.cell(row, 4, p["total_extra"])
            ws.cell(row, 5, p["total_debt"])
            for col, d in enumerate(shown_debts, 6):
                raw = mapping_get(p["balances"], d["id"], 0.0)
                if raw is None:
                    ws.cell(row, col, None)
                    continue
                bal = parse_amount(raw, allow_blank=True)
                if bal is None:
                    ws.cell(row, col, None)
                    continue
                if bal < 0.5:
                    cell = ws.cell(row, col, "✓")
                    cell.fill = PatternFill("solid", fgColor="A8E6CF")
                    cell.alignment = Alignment(horizontal="right")
                    cell.font = Font(bold=True, color="146C43")
                else:
                    ws.cell(row, col, bal)

        ws2 = wb.create_sheet("Debts")
        _dates, _overall, payoff_by_id = debt_free_dates_by_type(
            self.debts, self.settings, extras_by_period=self.extras_by_period,
        )
        debt_headers = ["ID", "Short Name", "Full Name", "Current", "Rate", "Minimum Payment", "Next Due", "Payment Frequency", "Debt Type", "Payoff Rank", "Estimated Payoff"]
        for col, h in enumerate(debt_headers, 1):
            cell = ws2.cell(1, col, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2C3E50")
        for row, d in enumerate(self.debts, 2):
            ws2.cell(row, 1, d["id"])
            ws2.cell(row, 2, d["short_name"])
            ws2.cell(row, 3, d["full_name"])
            ws2.cell(row, 4, d["current"])
            ws2.cell(row, 5, d["rate"])
            ws2.cell(row, 6, d["min_payment"])
            due_anchor, _due_frequency = due_series(d)
            ws2.cell(row, 7, format_display_date(due_anchor) if due_anchor else "")
            ws2.cell(row, 8, DUE_FREQUENCY_LABELS.get(d.get("due_frequency"), "") if d.get("due_md") else "")
            ws2.cell(row, 9, d.get("debt_type") or "")
            ws2.cell(row, 10, d.get("payoff_rank") or "")
            if not _is_active_debt(d):
                ws2.cell(row, 11, "✓")
            else:
                ws2.cell(row, 11, format_display_date(payoff_by_id.get(d.get("id"), "Never")))

        wb.save(path)
        self.status_var.set(f"Exported schedule to {os.path.basename(path)}")

    def export_archives(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing library", "Install openpyxl first:\n  pip install openpyxl")
            return
        if not self.archives:
            messagebox.showinfo("Archives", "No archived years yet.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        first = True
        for year in sorted(self.archives.keys()):
            if first:
                ws = wb.active
                ws.title = str(year)
                first = False
            else:
                ws = wb.create_sheet(str(year))
            ws.cell(1, 1, "Period").font = Font(bold=True)
            ws.cell(1, 2, "Date").font = Font(bold=True)
            ws.cell(1, 3, "Total Debt").font = Font(bold=True)
            for row, p in enumerate(self.archives[year], 2):
                ws.cell(row, 1, p["period"])
                ws.cell(row, 2, format_display_date(p.get("date")))
                ws.cell(row, 3, p["total_debt"])
        wb.save(path)
        self.status_var.set(f"Exported archives to {os.path.basename(path)}")

    def _select_tab(self, name):
        names = getattr(self.notebook, "_name_list", [])
        if name in names:
            self.notebook.set(name)

    def show_help(self):
        help_text = """Debt Payoff Planner

1. Settings
Set how often you are paid, the payday your balances belong to, and the extra you can add each payday. Yearly take-home pay is optional. It feeds the income and debt-to-income figures on Overview.

2. Debts
Add each balance, rate, minimum payment, and due day. Current Balance is what remains on the payday now in effect, after that day's minimum, interest, and any extra entered on the Pay Schedule. Rank 1 gets the extra payment first. Drag a row, or use Move Up and Move Down, to change that order. When a debt is paid off, its minimum joins the extra for the debts still open.

3. Pay Schedule
Each row is one payday. The column titles stay in place while you scroll. Double-click a row to record an extra for that day. Archive Completed Pay Periods files past paydays. Those rows stay under Archives and do not change later.

4. Overview, Charts, and Interest
Overview is the snapshot: balances, payoff dates, and debt-to-income. Debt by Debtor labels its payoff date as the date with estimated snowball payments. Charts and Interest compare extra payments in rank order with minimum payments only.

5. Saving
File → Save stores the plan and opens it next time. A sample plan is view only. Create your own plan, or use Save As, when you are ready to edit.

The program checks GitHub for a newer version when it opens. File → Check for Updates checks again.

Interest each payday is the balance times the annual rate, divided by the number of pay periods in a year.
"""
        messagebox.showinfo("How to Use", help_text)


class DebtDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Debt", debt=None):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()
        prepare = getattr(parent, "_prepare_dialog", None)
        if prepare:
            prepare(self)
        self.configure(fg_color=COLORS["bg"])
        self.grid_columnconfigure(1, weight=1, minsize=300)

        fields_before_type = [
            ("short_name", "Short Name", ""),
            ("full_name", "Full Name", ""),
            ("original", "Original Balance", "0"),
            ("current", "Current Balance", "0"),
            ("rate", "Annual Rate (e.g. 0.259 or 25.9%)", "0"),
            ("min_payment", "Minimum Payment", "0"),
        ]
        fields_after_type = [
            ("due_md", "Next Due (MM-DD)", ""),
            ("notes", "Notes", ""),
        ]
        self.vars = {}
        row = 0

        def place_field(widget, field_row):
            widget.configure(width=300)
            widget.grid(row=field_row, column=1, sticky="ew", padx=(0, 16), pady=6)

        for key, label, default in fields_before_type:
            top = 16 if row == 0 else 6
            AppLabel(self, text=label).grid(row=row, column=0, sticky="w", padx=(16, 10), pady=(top, 6))
            val = str(debt.get(key, default)) if debt else default
            var = tk.StringVar(value=val)
            entry = AppEntry(self, textvariable=var, width=32)
            place_field(entry, row)
            if row == 0:
                entry.grid_configure(pady=(16, 6))
            self.vars[key] = var
            row += 1

        AppLabel(self, text="Debt Type").grid(row=row, column=0, sticky="w", padx=(16, 10), pady=6)
        current_type = (debt or {}).get("debt_type") or ""
        if current_type not in DEBT_TYPES:
            current_type = ""
        self.type_var = tk.StringVar(value=current_type)
        type_combo = AppCombo(self, textvariable=self.type_var, values=DEBT_TYPES, state="readonly", width=29)
        place_field(type_combo, row)
        row += 1

        for key, label, default in fields_after_type:
            AppLabel(self, text=label).grid(row=row, column=0, sticky="w", padx=(16, 10), pady=6)
            if debt and key == "due_md":
                val = debt.get("due_md") or ""
                if not val:
                    parsed = parse_month_day(debt.get("due_date"))
                    val = f"{parsed[0]:02d}-{parsed[1]:02d}" if parsed else ""
            else:
                val = str(debt.get(key, default)) if debt else default
            var = tk.StringVar(value=val)
            entry = AppEntry(self, textvariable=var, width=32)
            place_field(entry, row)
            self.vars[key] = var
            row += 1

        freq_row = row
        AppLabel(self, text="Payment Frequency").grid(row=freq_row, column=0, sticky="w", padx=(16, 10), pady=6)
        current_freq = (debt or {}).get("due_frequency") or "monthly"
        self.freq_var = tk.StringVar(value=DUE_FREQUENCY_LABELS.get(current_freq, "Monthly"))
        freq_combo = AppCombo(
            self,
            textvariable=self.freq_var,
            values=[DUE_FREQUENCY_LABELS[key] for key in DUE_FREQUENCIES],
            state="readonly",
            width=29,
        )
        place_field(freq_combo, freq_row)

        self.include_paid_var = tk.BooleanVar(value=include_paid_minimum(debt) if debt else True)
        AppCheck(
            self,
            text="Include regular payment in Additional Per Pay when paid off",
            variable=self.include_paid_var,
        ).grid(row=freq_row + 1, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 2))

        btn = AppFrame(self)
        btn.grid(row=freq_row + 2, column=0, columnspan=2, pady=12)
        AppButton(btn, text="OK", command=self.on_ok).pack(side="left", padx=6)
        AppButton(btn, text="Cancel", style="secondary", command=self.destroy).pack(side="left", padx=6)

        self.wait_window()

    def on_ok(self):
        debt_type = self.type_var.get().strip()
        if debt_type not in DEBT_TYPES:
            messagebox.showerror("Error", "Choose a debt type.")
            return
        original = parse_amount(self.vars["original"].get(), minimum=0, maximum=1e12, allow_blank=True)
        current = parse_amount(self.vars["current"].get(), minimum=0, maximum=1e12, allow_blank=True)
        minimum = parse_amount(self.vars["min_payment"].get(), minimum=0, maximum=1e12, allow_blank=True)
        rate = parse_amount(self.vars["rate"].get(), minimum=0, maximum=1000, allow_blank=True)
        if None in (original, current, minimum, rate):
            messagebox.showerror("Error", "Balances, rate, and minimum must be non-negative numbers.")
            return
        if rate > 1:
            rate = rate / 100.0
        if rate > 10:
            messagebox.showerror("Error", "Annual rate is too large.")
            return
        due_raw = self.vars["due_md"].get().strip()
        parsed = parse_month_day(due_raw)
        if due_raw and not parsed:
            messagebox.showerror("Error", "Next due must be MM-DD, or left blank.")
            return
        frequency = next(
            (key for key, label in DUE_FREQUENCY_LABELS.items() if label == self.freq_var.get()),
            "monthly",
        )
        if parsed:
            month, day = parsed
            anchor = next_occurrence(month, day)
            due_md = f"{month:02d}-{day:02d}"
            due_anchor = anchor.isoformat()
        else:
            due_md = ""
            due_anchor = ""
        short_name = self.vars["short_name"].get().strip() or "Debt"
        self.result = {
            "short_name": short_name,
            "full_name": self.vars["full_name"].get().strip() or short_name,
            "original": original,
            "current": current,
            "rate": rate,
            "min_payment": minimum,
            "debt_type": debt_type,
            "due_md": due_md,
            "due_anchor": due_anchor,
            "due_date": due_anchor,
            "due_frequency": frequency,
            "notes": self.vars["notes"].get().strip(),
            "include_paid_minimum": bool(self.include_paid_var.get()),
        }
        self.destroy()


def _monitor_center(width, height):
    """Screen position that centers a window on the monitor under the cursor."""
    try:
        import ctypes
        user32 = ctypes.windll.user32

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("rcMonitor", RECT),
                ("rcWork", RECT),
                ("dwFlags", ctypes.c_ulong),
            ]

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
        user32.MonitorFromPoint.argtypes = [POINT, ctypes.c_ulong]
        user32.MonitorFromPoint.restype = ctypes.c_void_p
        user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.POINTER(MONITORINFO)]
        user32.GetMonitorInfoW.restype = ctypes.c_int

        point = POINT()
        user32.GetCursorPos(ctypes.byref(point))
        monitor = user32.MonitorFromPoint(point, 2)
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(MONITORINFO)
        if monitor and user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            area = info.rcMonitor
            x = area.left + max(0, ((area.right - area.left) - width) // 2)
            y = area.top + max(0, ((area.bottom - area.top) - height) // 2)
            return x, y, area.left, area.top, area.right, area.bottom
    except Exception:
        return None
    return None


def _place_splash(splash, width, height):
    origin = _monitor_center(width, height)
    if origin is None:
        screen_w = max(splash.winfo_screenwidth(), width)
        screen_h = max(splash.winfo_screenheight(), height)
        splash.geometry(f"{width}x{height}+{(screen_w - width) // 2}+{(screen_h - height) // 2}")
        return
    x, y = origin[0], origin[1]
    splash.geometry(f"{width}x{height}+{x}+{y}")


def _fade_splash(splash, start, end, duration):
    started = time.perf_counter()
    while True:
        elapsed = time.perf_counter() - started
        portion = 1.0 if duration <= 0 or elapsed >= duration else elapsed / duration
        try:
            splash.attributes("-alpha", start + (end - start) * portion)
            splash.update()
        except tk.TclError:
            return
        if elapsed >= duration:
            return
        time.sleep(0.016)


def show_opening_splash(root):
    """Fade the splash in, centered, and leave it up while the planner loads."""
    scale = 1.0
    try:
        from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker
        scale = max(1.0, float(ScalingTracker.get_window_dpi_scaling(root)))
    except Exception:
        scale = 1.0
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg=COLORS["header_bg"])
    try:
        splash.attributes("-topmost", True)
        splash.attributes("-alpha", 0.0)
    except tk.TclError:
        pass
    width = int(round(520 * scale))
    height = int(round(220 * scale))
    tk.Label(
        splash,
        text=APP_NAME,
        bg=COLORS["header_bg"],
        fg=COLORS["header_fg"],
        font=ui_font( max(22, int(round(26 * scale))), "bold"),
    ).pack(expand=True, pady=(int(round(28 * scale)), 0))
    tk.Label(
        splash,
        text=f"Version {APP_VERSION}",
        bg=COLORS["header_bg"],
        fg=COLORS["header_fg"],
        font=ui_font( max(12, int(round(13 * scale)))),
    ).pack(pady=(0, int(round(36 * scale))))
    splash.update_idletasks()
    _place_splash(splash, width, height)
    try:
        splash.deiconify()
        splash.lift()
        splash.update()
    except tk.TclError:
        return splash
    _place_splash(splash, width, height)
    _fade_splash(splash, 0.0, 1.0, 0.35)
    _place_splash(splash, width, height)
    try:
        splash.attributes("-topmost", True)
        splash.lift()
    except tk.TclError:
        pass
    return splash


def close_opening_splash(splash):
    """Fade the splash out once the planner window is open."""
    if splash is None:
        return
    try:
        if not splash.winfo_exists():
            return
        splash.attributes("-topmost", True)
        splash.lift()
    except tk.TclError:
        return
    _fade_splash(splash, 1.0, 0.0, 0.45)
    try:
        splash.destroy()
    except tk.TclError:
        pass


def app_icon_path():
    """Icon used for the window, taskbar, desktop shortcut, and File Explorer."""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(getattr(sys, "_MEIPASS", ""), "debt_payoff.ico"))
        candidates.append(os.path.join(os.path.dirname(sys.executable), "debt_payoff.ico"))
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "debt_payoff.ico"))
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def app_png_path():
    """PNG icon. Linux window icons cannot use the Windows .ico file."""
    candidates = [os.path.join(os.path.dirname(os.path.abspath(__file__)), "debt_payoff.png")]
    ico = app_icon_path()
    if ico:
        candidates.append(os.path.splitext(ico)[0] + ".png")
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def apply_window_icon(window):
    if window is None:
        return
    if not sys.platform.startswith("win"):
        png = app_png_path()
        if not png:
            return
        try:
            image = tk.PhotoImage(file=png)
            window.iconphoto(True, image)
            window._icon_image = image
        except tk.TclError:
            return
        return
    path = app_icon_path()
    if not path:
        return
    try:
        window.iconbitmap(path)
        window.iconbitmap(default=path)
    except tk.TclError:
        return


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DebtPayoffPlanner.DebtPayoffPlanner")
        except Exception:
            pass
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    choose_ui_font_family(root)
    apply_window_icon(root)
    root.withdraw()
    splash = None
    try:
        splash = show_opening_splash(root)
        app = DebtPayoffApp(root)
        try:
            root.update()
        except tk.TclError:
            pass
    finally:
        close_opening_splash(splash)
    root.after(200, app._startup_prompts)
    root.mainloop()


if __name__ == "__main__":
    main()
