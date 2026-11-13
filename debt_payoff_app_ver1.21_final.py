#!/usr/bin/env python3
"""
Debt Payoff Planner - Desktop App
Standalone Tkinter application for bi-weekly / semi-monthly / monthly debt payoff planning
with snowball extras, year archiving, and Excel export.
Styled with professional colors matching the original Excel layout.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import math
import datetime
from datetime import timedelta
from copy import deepcopy
import os
import sys

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
}

# Professional color palette (no bright/fluorescent)
COLORS = {
    "bg": "#f0f2f5",
    "header_bg": "#1a365d",
    "header_fg": "#ffffff",
    "green": "#a8e6cf",
    "green_alt": "#8ed4b8",
    "paid_row": "#86d07c",
    "paid_row_alt": "#5aaa58",
    "blue": "#aed6f1",
    "alt_row": "#e8eef4",
    "border": "#bdc3c7",
    "text": "#1a1a1a",
    "accent": "#2980b9",
}

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


def chart_mix_white(color, weight):
    """Lighten a hex color toward white. weight 0 keeps the color, 1 is white."""
    value = str(color).lstrip("#")
    if len(value) != 6:
        return color
    channels = [int(value[index:index + 2], 16) for index in (0, 2, 4)]
    mixed = [round(channel + (255 - channel) * weight) for channel in channels]
    return "#{:02x}{:02x}{:02x}".format(*mixed)


def chart_tick_label(day, detailed):
    """Month and 'YY, or a pay date when the span is under a year."""
    if detailed:
        return day.strftime("%m/%d/%y")
    return f"{MONTH_NAMES[day.month - 1]} '{day.year % 100:02d}"


def chart_axis_ticks(start, end, frequency, plot_width, detailed=None):
    """Tick positions and labels that fit the chart width."""
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
    sample = items[0][1]
    slot = max(58, int(len(sample) * (6.4 if detailed else 4.4)))
    max_labels = max(2, int(max(int(plot_width), slot) / slot))
    stride = max(1, math.ceil(len(items) / max_labels))
    chosen = list(items[::stride])
    if items[-1] not in chosen:
        chosen.append(items[-1])
    return detailed, chosen


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


def chart_series(points, debts, mode, selection=None):
    """Balance lines for every type or debtor, or for one selected name."""
    active = _active_debts_in_order(debts)
    groups = []
    if mode == "debtor":
        for index, debt in enumerate(active):
            name = debt.get("short_name") or f"Debt {debt.get('id')}"
            if selection and name != selection:
                continue
            groups.append((name, index, [debt["id"]]))
        title = selection or "All Debtors"
    else:
        by_type = {}
        for debt in active:
            by_type.setdefault(_debt_type_name(debt), []).append(debt["id"])
        for index, name in enumerate(OVERVIEW_TYPES):
            ids = by_type.get(name)
            if not ids:
                continue
            if selection and name != selection:
                continue
            groups.append((name, index, ids))
        title = selection or "All Debt Types"
    series = []
    for name, color_index, ids in groups:
        plotted = []
        for day, balances in points:
            total = sum(max(0.0, balances.get(did, 0.0)) for did in ids)
            plotted.append((day, total))
            if total < 0.5 and len(plotted) > 1:
                break
        if plotted:
            series.append({
                "name": name,
                "color": CHART_COLORS[color_index % len(CHART_COLORS)],
                "points": plotted,
            })
    return title, series


def chart_debt_lines(points, debts, type_name=None):
    """One balance line per debt. type_name limits the list to that debt type."""
    active = _active_debts_in_order(debts)
    order = {debt["id"]: index for index, debt in enumerate(active)}
    chosen = active
    if type_name:
        chosen = [debt for debt in active if _debt_type_name(debt) == type_name]
    series = []
    for debt in chosen:
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


def chart_window_end(series, fallback):
    """Right edge is the last visible debt that reaches zero."""
    ends = []
    for item in series:
        points = item.get("points") or []
        if points and points[-1][1] < 0.5:
            ends.append(points[-1][0])
    if ends:
        return max(ends)
    last_dates = [item["points"][-1][0] for item in series if item.get("points")]
    return max(last_dates) if last_dates else fallback


def chart_time_scale(start, series, viewport_width):
    """Widen the plot when one line runs far past the rest.

    The visible window stays scaled to the earlier payoffs. The longer line
    keeps that same scale, so the date labels stay closer and the rest is
    reached by scrolling.
    """
    viewport_width = max(int(viewport_width), 40)
    if start is None:
        return start, 1, False, viewport_width
    spans = []
    full_end = start
    for item in series or []:
        points = item.get("points") or []
        if not points:
            continue
        day = points[-1][0]
        if day > full_end:
            full_end = day
        spans.append(max((day - start).days, 0))
    if not spans:
        return full_end, 1, False, viewport_width
    unique = sorted(set(spans))
    bulk_days = unique[-1]
    for index in range(len(unique) - 1):
        current = unique[index]
        nxt = unique[index + 1]
        if current >= 14 and nxt >= current * 2 and (nxt - current) >= 365:
            bulk_days = current
            break
    full_days = max((full_end - start).days, 1)
    bulk_days = max(min(bulk_days, full_days), 1)
    if full_days <= bulk_days:
        return full_end, full_days, False, viewport_width
    plot_width = int(math.ceil(viewport_width * (full_days / bulk_days)))
    plot_width = min(max(plot_width, viewport_width + 1), 20000)
    if plot_width <= viewport_width:
        return full_end, full_days, False, viewport_width
    return full_end, bulk_days, True, plot_width


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

# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class DebtPayoffApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Debt Payoff Planner")
        self.root.geometry("1400x800")
        self.root.minsize(1100, 650)
        try:
            self.root.state("zoomed")
        except tk.TclError:
            pass
        self.root.configure(bg=COLORS["bg"])

        self.debts = []
        self.settings = deepcopy(DEFAULT_SETTINGS)
        self.extras_by_period = {}
        self.archives = {}
        self.current_file = None
        self.schedule = []
        self.undo_stack = []  # up to 15 snapshots
        self.redo_stack = []
        self._settings_lock = False

        self._setup_styles()
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
        self.root.after(200, self._startup_prompts)

    def _prepare_dialog(self, dialog):
        """Keep a dialog with the main window so minimizing cannot leave it stuck."""
        dialog.transient(self.root)
        try:
            dialog.attributes("-toolwindow", True)
        except tk.TclError:
            pass

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

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"), foreground=COLORS["header_bg"])
        style.configure("Overview.TLabelframe.Label", background=COLORS["bg"], foreground=COLORS["header_bg"], font=("Segoe UI", 16, "bold"))
        style.configure("OverviewCaption.TLabel", background=COLORS["bg"], foreground="#5a6a7a", font=("Segoe UI", 9))
        style.configure("OverviewValue.TLabel", background=COLORS["bg"], foreground=COLORS["header_bg"], font=("Segoe UI", 11, "bold"))
        style.configure("OverviewHero.TLabel", background=COLORS["bg"], foreground=COLORS["header_bg"], font=("Segoe UI", 16, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.map("TButton", background=[("active", COLORS["accent"])])
        style.configure("TNotebook", background=COLORS["bg"], bordercolor=COLORS["border"], tabmargins=[2, 2, 2, 0])
        # Clam keeps a smaller padding on the selected tab. Swap that so the
        # open tab uses the larger size and the others use the smaller one.
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding="4.5p 3p 4.5p 1.5p",
            background="#d9dee6",
            lightcolor="#e7ecf1",
            darkcolor="#c5ced8",
            bordercolor=COLORS["border"],
        )
        style.map(
            "TNotebook.Tab",
            padding=[("selected", [12, 6, 12, 6])],
            background=[("selected", COLORS["bg"])],
            lightcolor=[("selected", COLORS["bg"])],
            darkcolor=[("selected", "#d5dbe3")],
        )
        style.configure(
            "Overview.TLabelframe",
            background=COLORS["bg"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["bg"],
            darkcolor="#d5dbe3",
            relief="groove",
            borderwidth=1,
        )
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=22, background=COLORS["bg"], fieldbackground=COLORS["bg"])
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background=COLORS["header_bg"], foreground=COLORS["header_fg"])
        style.map("Treeview", background=[("selected", COLORS["blue"])])

    def _build_menu(self):
        menubar = tk.Menu(self.root, bg=COLORS["header_bg"], fg=COLORS["header_fg"])
        file_menu = tk.Menu(menubar, tearoff=0)
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
        file_menu.add_command(label="Exit", command=self.confirm_close)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Tutorial...", command=self.start_tutorial)
        help_menu.add_command(label="How to Use", command=self.show_help)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.notebook = notebook

        overview_frame = ttk.Frame(notebook)
        notebook.add(overview_frame, text="Overview")
        self._build_overview_tab(overview_frame)

        schedule_frame = ttk.Frame(notebook)
        notebook.add(schedule_frame, text="Pay Schedule")
        self._build_schedule_tab(schedule_frame)

        debts_frame = ttk.Frame(notebook)
        notebook.add(debts_frame, text="Debts")
        self._build_debts_tab(debts_frame)

        charts_frame = ttk.Frame(notebook)
        notebook.add(charts_frame, text="Charts")
        self._build_charts_tab(charts_frame)

        archives_frame = ttk.Frame(notebook)
        notebook.add(archives_frame, text="Archives")
        self._build_archives_tab(archives_frame)

        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="Settings")
        self._build_settings_tab(settings_frame)

        self.snowball_var = tk.StringVar(value="$0.00")
        snowball = tk.Toplevel(self.root)
        snowball.overrideredirect(True)
        snowball.configure(bg="black")
        snowball.withdraw()
        try:
            snowball.transient(self.root)
            snowball.attributes("-toolwindow", True)
        except tk.TclError:
            pass
        tk.Label(snowball, text="Additional Per Pay", bg="black", fg="#d5d5d5", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(10, 8))
        tk.Label(snowball, textvariable=self.snowball_var, bg="black", fg="#3DDC67", font=("Segoe UI", 13, "bold")).pack(side="left", padx=(0, 10))
        self._snowball = snowball
        notebook.bind("<Configure>", self._align_snowball, add="+")
        self.root.bind("<Configure>", self._align_snowball, add="+")
        self.root.after_idle(self._align_snowball)

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", background=COLORS["header_bg"], foreground=COLORS["header_fg"])
        status.pack(side="bottom", fill="x")

    def _selected_tab_box(self):
        notebook = self.notebook
        try:
            index = notebook.index(notebook.select())
        except tk.TclError:
            return None
        width = notebook.winfo_width()
        limit = notebook.nametowidget(notebook.select()).winfo_y() + 6
        hit_x = None
        for x in range(4, max(width, 4), 8):
            try:
                if notebook.index(f"@{x},8") == index:
                    hit_x = x
                    break
            except tk.TclError:
                continue
        if hit_x is None:
            return None
        top = bottom = None
        for y in range(0, max(limit, 1)):
            try:
                hit = notebook.index(f"@{hit_x},{y}")
            except tk.TclError:
                hit = None
            if hit == index:
                if top is None:
                    top = y
                bottom = y
            elif top is not None:
                break
        if top is None or bottom is None:
            return None
        return top, bottom - top + 1

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
            if not notebook.winfo_exists() or not notebook.select():
                return
            notebook.update_idletasks()
            box = self._selected_tab_box()
            if box is None:
                return
            top, height = box
            if height < 8:
                return
            snowball.update_idletasks()
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
        except tk.TclError:
            return
        finally:
            self._aligning_snowball = False

    def _build_debts_tab(self, parent):
        cols = ("short", "full", "current", "rate", "min", "due", "every", "type", "rank", "est")
        self.debt_tree = ttk.Treeview(parent, columns=cols, show="headings", height=16)
        self.debt_tree.heading("short", text="Short Name", anchor="w")
        self.debt_tree.heading("full", text="Full Name", anchor="w")
        self.debt_tree.heading("current", text="Current Balance", anchor="e")
        self.debt_tree.heading("rate", text="Rate", anchor="e")
        self.debt_tree.heading("min", text="Minimum Payment", anchor="e")
        self.debt_tree.heading("due", text="Next Due", anchor="center")
        self.debt_tree.heading("every", text="Payment Frequency", anchor="center")
        self.debt_tree.heading("type", text="Debt Type", anchor="center")
        self.debt_tree.heading("rank", text="Payoff Rank", anchor="center")
        self.debt_tree.heading("est", text="Estimated Payoff", anchor="center")
        self.debt_tree.column("short", width=110, anchor="w")
        self.debt_tree.column("full", width=190, anchor="w")
        self.debt_tree.column("current", width=120, anchor="e")
        self.debt_tree.column("rate", width=70, anchor="e")
        self.debt_tree.column("min", width=120, anchor="e")
        self.debt_tree.column("due", width=110, anchor="center")
        self.debt_tree.column("every", width=130, anchor="center")
        self.debt_tree.column("type", width=110, anchor="center")
        self.debt_tree.column("rank", width=90, anchor="center")
        self.debt_tree.column("est", width=120, anchor="center")
        self.debt_tree.pack(fill="both", expand=True, padx=8, pady=8)
        self._debt_drag = None
        self.debt_tree.bind("<ButtonPress-1>", self._on_debt_drag_press, add="+")
        self.debt_tree.bind("<B1-Motion>", self._on_debt_drag_motion, add="+")
        self.debt_tree.bind("<ButtonRelease-1>", self._on_debt_drag_release, add="+")

        self.debt_tree.tag_configure("zero", background=COLORS["paid_row"])
        self.debt_tree.tag_configure("zero_alt", background=COLORS["paid_row_alt"])
        self.debt_tree.tag_configure("alt", background=COLORS["alt_row"])

        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=8, pady=6)
        self.debt_btn_frame = btn_frame
        ttk.Button(btn_frame, text="Add Debt", command=self.add_debt).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Edit Debt", command=self.edit_debt).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Delete Debt", command=self.delete_debt).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Update Current Balances", command=self.update_balances).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Move Down", command=lambda: self.move_debt(1)).pack(side="right", padx=3)
        ttk.Button(btn_frame, text="Move Up", command=lambda: self.move_debt(-1)).pack(side="right", padx=3)

        self.refresh_debt_tree()

    def _build_settings_tab(self, parent):
        frm = ttk.Frame(parent, padding=20)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="First Pay Date (MM-DD-YYYY):", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=6)
        self.date_var = tk.StringVar(value=self.settings["first_pay_date"])
        date_entry = ttk.Entry(frm, textvariable=self.date_var, width=24)
        date_entry.grid(row=0, column=1, sticky="w", pady=6)
        date_entry.bind("<FocusOut>", self._commit_settings)
        date_entry.bind("<Return>", self._commit_settings)
        self.settings_first_pay = date_entry

        ttk.Label(frm, text="Balances As Of (MM-DD-YYYY):", style="Header.TLabel").grid(row=1, column=0, sticky="w", pady=6)
        self.anchor_date_var = tk.StringVar(value=self.settings.get("balance_as_of") or self.settings["first_pay_date"])
        anchor_entry = ttk.Entry(frm, textvariable=self.anchor_date_var, width=24)
        anchor_entry.grid(row=1, column=1, sticky="w", pady=6)
        anchor_entry.bind("<FocusOut>", self._commit_settings)
        anchor_entry.bind("<Return>", self._commit_settings)
        self.settings_balance_as_of = anchor_entry

        ttk.Label(frm, text="Starting Additional Per Pay ($):", style="Header.TLabel").grid(row=2, column=0, sticky="w", pady=6)
        self.starting_var = tk.StringVar(value=str(self.settings.get("starting_additional", 0)))
        starting_entry = ttk.Entry(frm, textvariable=self.starting_var, width=24)
        starting_entry.grid(row=2, column=1, sticky="w", pady=6)
        starting_entry.bind("<FocusOut>", self._commit_settings)
        starting_entry.bind("<Return>", self._commit_settings)
        self.settings_starting = starting_entry

        ttk.Label(frm, text="Projection Periods:", style="Header.TLabel").grid(row=3, column=0, sticky="w", pady=6)
        self.proj_var = tk.StringVar(value=str(self.settings["projection_periods"]))
        proj_entry = ttk.Entry(frm, textvariable=self.proj_var, width=24)
        proj_entry.grid(row=3, column=1, sticky="w", pady=6)
        proj_entry.bind("<FocusOut>", self._commit_settings)
        proj_entry.bind("<Return>", self._commit_settings)

        ttk.Label(frm, text="Yearly Net Salary ($):", style="Header.TLabel").grid(row=4, column=0, sticky="w", pady=6)
        salary_row = ttk.Frame(frm)
        salary_row.grid(row=4, column=1, sticky="w", pady=6)
        self.salary_var = tk.StringVar(value=str(self.settings.get("yearly_net_salary", 0)))
        salary_entry = ttk.Entry(salary_row, textvariable=self.salary_var, width=16)
        salary_entry.pack(side="left")
        salary_entry.bind("<FocusOut>", self._commit_settings)
        salary_entry.bind("<Return>", self._commit_settings)
        self.settings_salary = salary_entry
        ttk.Label(salary_row, text="Pay Frequency").pack(side="left", padx=(8, 0))
        self.freq_var = tk.StringVar(value=self.settings["frequency"])
        freq_combo = ttk.Combobox(
            salary_row, textvariable=self.freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        freq_combo.pack(side="left", padx=(8, 0))
        freq_combo.bind("<<ComboboxSelected>>", self._commit_settings)
        self.settings_freq = freq_combo

        ttk.Label(frm, text="Spouse / Secondary Yearly Net Income ($):", style="Header.TLabel").grid(row=5, column=0, sticky="w", pady=6)
        spouse_row = ttk.Frame(frm)
        spouse_row.grid(row=5, column=1, sticky="w", pady=6)
        self.spouse_var = tk.StringVar(value=str(self.settings.get("spouse_salary", 0)))
        spouse_entry = ttk.Entry(spouse_row, textvariable=self.spouse_var, width=16)
        spouse_entry.pack(side="left")
        spouse_entry.bind("<FocusOut>", self._commit_settings)
        spouse_entry.bind("<Return>", self._commit_settings)
        self.settings_spouse = spouse_entry
        ttk.Label(spouse_row, text="Pay Frequency").pack(side="left", padx=(8, 0))
        spouse_freq_value = self.settings.get("spouse_income_frequency", "biweekly")
        if spouse_freq_value not in PAY_FREQUENCIES:
            spouse_freq_value = "biweekly"
        self.spouse_freq_var = tk.StringVar(value=spouse_freq_value)
        spouse_freq = ttk.Combobox(
            spouse_row, textvariable=self.spouse_freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        spouse_freq.pack(side="left", padx=(8, 0))
        spouse_freq.bind("<<ComboboxSelected>>", self._commit_settings)

        ttk.Label(frm, text="Other Yearly Net Salary ($):", style="Header.TLabel").grid(row=6, column=0, sticky="w", pady=6)
        other_row = ttk.Frame(frm)
        other_row.grid(row=6, column=1, sticky="w", pady=6)
        self.other_income_var = tk.StringVar(value=str(self.settings.get("other_income", 0)))
        other_entry = ttk.Entry(other_row, textvariable=self.other_income_var, width=16)
        other_entry.pack(side="left")
        other_entry.bind("<FocusOut>", self._commit_settings)
        other_entry.bind("<Return>", self._commit_settings)
        self.settings_other = other_entry
        ttk.Label(other_row, text="Pay Frequency").pack(side="left", padx=(8, 0))
        other_freq_value = self.settings.get("other_income_frequency", "biweekly")
        if other_freq_value not in PAY_FREQUENCIES:
            other_freq_value = "biweekly"
        self.other_freq_var = tk.StringVar(value=other_freq_value)
        other_freq = ttk.Combobox(
            other_row, textvariable=self.other_freq_var,
            values=list(PAY_FREQUENCIES),
            state="readonly", width=14,
        )
        other_freq.pack(side="left", padx=(8, 0))
        other_freq.bind("<<ComboboxSelected>>", self._commit_settings)

        note = ttk.Label(frm, text=(
            "Notes:\n"
            "• A completed entry is applied when you leave the field. Invalid entries are left unchanged.\n"
            "• Yearly Net Salary, Spouse / Secondary Yearly Net Income, and Other Yearly Net Salary are yearly net amounts. The spouse and other dropdowns only calculate Overview Pay Period Average from that yearly figure. Monthly Average is yearly net ÷ 12. Totals and DTI use the yearly figures.\n"
            "• Starting Additional Per Pay is the base snowball. A paid-off debt adds its regular payment unless that debt is set to leave it out.\n"
            "• Balances As Of is the pay date your current balances belong to. Changing First Pay Date does not move them.\n"
            "• Each debt's minimum payment is subtracted on the pay period that contains its due date.\n"
            "• Payment Frequency on the debt controls how often that minimum repeats: weekly, bi-weekly, semi-monthly, or monthly.\n"
            "• Debt Type is Auto, Personal, Mortgage, Credit, or Payment Plan.\n"
            "• Payoff Rank is the order extra payments go. Drag a row or use Move Up / Move Down. A new debt is last. Estimated Payoff on that tab uses this order plus Starting Additional Per Pay.\n"
            "• When a debt is paid off it is removed from ranking, moved to the bottom, and dropped from remaining-debt estimates. The others keep their relative order and are numbered 1, 2, 3 again.\n"
            "• Pay periods before the balance date are estimated from the balance, rate, and minimum payment.\n"
            "• A newly added debt starts on the current pay period. Earlier periods use the paid-off placeholder.\n"
            "• The period number counts from the first pay date of each calendar year.\n"
            "• Archive removes completed periods. A later change does not put them back. Deleting a debt leaves its archived periods as they were.\n"
            "• A debt that is paid off drops off the pay schedule when you archive, and also the next time the plan is opened. It stays on the Debts tab until you delete it. Give it a balance again and it returns on the current pay period.\n"
            "• Double-click a pay-schedule cell to assign an extra for that debt and period.\n"
            "• Green cells are paid-off balances.\n"
            "• File → Save writes the open plan to the JSON file."
        ), justify="left")
        note.grid(row=7, column=0, columnspan=2, sticky="w", pady=10)

    def _build_schedule_tab(self, parent):
        # Frozen Period/Pay Date plus a separately scrolling money grid
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True, padx=8, pady=8)

        self.sched_frozen_canvas = tk.Canvas(container, bg="white", highlightthickness=0, width=170)
        self.sched_canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical")
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.sched_canvas.xview)

        def _sched_yview(*args):
            self.sched_frozen_canvas.yview(*args)
            self.sched_canvas.yview(*args)

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

        self.sched_frozen_inner = ttk.Frame(self.sched_frozen_canvas)
        self.sched_inner = ttk.Frame(self.sched_canvas)
        self.sched_frozen_canvas.create_window((0, 0), window=self.sched_frozen_inner, anchor="nw")
        self.sched_canvas.create_window((0, 0), window=self.sched_inner, anchor="nw")

        def _fit_schedule_panes(_event=None):
            self.sched_frozen_inner.update_idletasks()
            self.sched_inner.update_idletasks()
            frozen_box = self.sched_frozen_canvas.bbox("all")
            body_box = self.sched_canvas.bbox("all")
            if frozen_box:
                self.sched_frozen_canvas.configure(scrollregion=frozen_box, width=max(frozen_box[2], 1))
            if body_box:
                self.sched_canvas.configure(scrollregion=body_box)

        self.sched_frozen_inner.bind("<Configure>", _fit_schedule_panes)
        self.sched_inner.bind("<Configure>", _fit_schedule_panes)
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
                    self.archive_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return
                if widget in schedule_parts:
                    _sched_yview("scroll", int(-1 * (event.delta / 120)), "units")
                    return
                widget = widget.master

        self.sched_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=8, pady=6)
        ttk.Label(btn_frame, text="Double-click a row to assign extras for that period").pack(side="left")
        archive_btn = ttk.Button(btn_frame, text="Archive Completed Pay Periods", command=self.archive_years)
        archive_btn.pack(side="right", padx=3)
        self.sched_archive_btn = archive_btn
        self.sched_btn_frame = btn_frame

    def _build_archives_tab(self, parent):
        year_bar = ttk.Frame(parent)
        year_bar.pack(fill="x", padx=8, pady=(10, 0))
        self.archive_year_bar = year_bar
        self._archive_year = None

        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True, padx=8, pady=8)
        self.archive_canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.archive_canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.archive_canvas.xview)
        self.archive_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.archive_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.archive_inner = ttk.Frame(self.archive_canvas)
        self.archive_canvas.create_window((0, 0), window=self.archive_inner, anchor="nw")
        self.archive_inner.bind("<Configure>", lambda e: self.archive_canvas.configure(scrollregion=self.archive_canvas.bbox("all")))
        self.refresh_archives()

    def _build_overview_tab(self, parent):
        frm = ttk.Frame(parent, padding=16)
        frm.pack(fill="both", expand=True)

        cards = ttk.Frame(frm)
        cards.grid(row=0, column=0, sticky="nsew")
        cards.grid_columnconfigure(0, weight=1)
        cards.grid_columnconfigure(1, weight=1)
        cards.grid_columnconfigure(2, weight=1)

        income = ttk.LabelFrame(cards, text="Income", style="Overview.TLabelframe", padding=10)
        income.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        value_pad = (16, 0)
        ttk.Label(income, text="Yearly", style="OverviewValue.TLabel", anchor="e").grid(row=0, column=1, sticky="ew", padx=value_pad)
        ttk.Label(
            income, text="Pay Period Average", style="OverviewValue.TLabel",
            anchor="e", justify="right", wraplength=140,
        ).grid(row=0, column=2, sticky="ew", padx=value_pad)
        ttk.Label(
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
        ttk.Label(income, text="Yearly Net Salary", style="OverviewValue.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Label(income, textvariable=self.overview_salary_year, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=1, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_salary_pay, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=2, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_salary_month, style="OverviewValue.TLabel", anchor="e").grid(row=1, column=3, sticky="ew", padx=value_pad)
        ttk.Label(income, text="Spouse / Secondary Yearly Net Income", style="OverviewValue.TLabel").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Label(income, textvariable=self.overview_spouse_year, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=1, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_spouse_pay, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=2, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_spouse_month, style="OverviewValue.TLabel", anchor="e").grid(row=2, column=3, sticky="ew", padx=value_pad)
        ttk.Label(income, text="Other Yearly Net Salary", style="OverviewValue.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Label(income, textvariable=self.overview_other_year, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=1, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_other_pay, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=2, sticky="ew", padx=value_pad)
        ttk.Label(income, textvariable=self.overview_other_month, style="OverviewValue.TLabel", anchor="e").grid(row=3, column=3, sticky="ew", padx=value_pad)
        ttk.Label(income, text="Total", style="OverviewValue.TLabel").grid(row=4, column=0, sticky="w", pady=(12, 2))
        ttk.Label(income, textvariable=self.overview_yearly_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=1, sticky="ew", padx=value_pad, pady=(12, 2))
        ttk.Label(income, textvariable=self.overview_pay_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=2, sticky="ew", padx=value_pad, pady=(12, 2))
        ttk.Label(income, textvariable=self.overview_monthly_income, style="OverviewHero.TLabel", anchor="e").grid(row=4, column=3, sticky="ew", padx=value_pad, pady=(12, 2))
        income.grid_columnconfigure(0, weight=1)
        for column in (1, 2, 3):
            income.grid_columnconfigure(column, weight=1, uniform="income_value")
        self.overview_income = income

        dti = ttk.LabelFrame(cards, text="Debt-to-Income", style="Overview.TLabelframe", padding=12)
        dti.grid(row=0, column=1, sticky="nsew", padx=4)
        self.overview_front_dti = tk.StringVar(value="—")
        self.overview_back_dti = tk.StringVar(value="—")
        ttk.Label(dti, text="Front-End", style="OverviewValue.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(dti, text="Mortgage Only", style="OverviewCaption.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 2))
        ttk.Label(dti, textvariable=self.overview_front_dti, style="OverviewHero.TLabel").grid(row=0, column=1, rowspan=2, sticky="e", padx=8)
        ttk.Label(dti, text="Back-End", style="OverviewValue.TLabel").grid(row=2, column=0, sticky="w", pady=(14, 0))
        ttk.Label(dti, text="All Active Debts", style="OverviewCaption.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 2))
        ttk.Label(dti, textvariable=self.overview_back_dti, style="OverviewHero.TLabel").grid(row=2, column=1, rowspan=2, sticky="e", padx=8, pady=(14, 0))
        dti.grid_columnconfigure(0, weight=1)
        self.overview_dti = dti

        mins = ttk.LabelFrame(cards, text="Minimum Payments", style="Overview.TLabelframe", padding=12)
        mins.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        self.overview_monthly_min = tk.StringVar(value="$0.00")
        self.overview_per_pay = tk.StringVar(value="$0.00")
        self.overview_per_pay_label = tk.StringVar(value="Per Pay Period")
        ttk.Label(mins, text="Monthly", style="OverviewValue.TLabel").grid(row=0, column=0, sticky="w", pady=10)
        ttk.Label(mins, textvariable=self.overview_monthly_min, style="OverviewHero.TLabel").grid(row=0, column=1, sticky="e", pady=10)
        ttk.Label(mins, textvariable=self.overview_per_pay_label, style="OverviewValue.TLabel").grid(row=1, column=0, sticky="w", pady=10)
        ttk.Label(mins, textvariable=self.overview_per_pay, style="OverviewHero.TLabel").grid(row=1, column=1, sticky="e", pady=10)
        mins.grid_columnconfigure(0, weight=1)
        self.overview_mins = mins

        debtor = ttk.LabelFrame(frm, text="By Debtor", style="Overview.TLabelframe", padding=10)
        debtor.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        debtor_wrap = ttk.Frame(debtor)
        debtor_wrap.pack(fill="both", expand=True)
        debtor_cols = ("name", "remaining", "monthly", "per_pay", "free")
        self.overview_debtor_tree = ttk.Treeview(debtor_wrap, columns=debtor_cols, show="headings", height=6)
        self.overview_debtor_tree.heading("name", text="Debtor", anchor="w")
        self.overview_debtor_tree.heading("remaining", text="Remaining", anchor="e")
        self.overview_debtor_tree.heading("monthly", text="Monthly Minimum", anchor="e")
        self.overview_debtor_tree.heading("per_pay", text="Per Pay Period", anchor="e")
        self.overview_debtor_tree.heading("free", text="Debt Free Estimated Date", anchor="center")
        self.overview_debtor_tree.column("name", width=160, anchor="w")
        self.overview_debtor_tree.column("remaining", width=120, anchor="e")
        self.overview_debtor_tree.column("monthly", width=150, anchor="e")
        self.overview_debtor_tree.column("per_pay", width=130, anchor="e")
        self.overview_debtor_tree.column("free", width=190, anchor="center")
        debtor_scroll = ttk.Scrollbar(debtor_wrap, orient="vertical", command=self.overview_debtor_tree.yview)
        self.overview_debtor_tree.configure(yscrollcommand=debtor_scroll.set)
        self.overview_debtor_tree.pack(side="left", fill="both", expand=True)
        debtor_scroll.pack(side="right", fill="y")
        self.overview_debtor_tree.tag_configure("alt", background=COLORS["alt_row"])
        self.overview_debtor_tree.tag_configure("total", background=COLORS["blue"])
        self.overview_debtor = debtor

        types = ttk.LabelFrame(frm, text="Debt by Type", style="Overview.TLabelframe", padding=10)
        types.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        cols = ("type", "remaining", "monthly", "per_pay", "free")
        self.overview_tree = ttk.Treeview(types, columns=cols, show="headings", height=8)
        self.overview_tree.heading("type", text="Debt Type", anchor="w")
        self.overview_tree.heading("remaining", text="Remaining", anchor="e")
        self.overview_tree.heading("monthly", text="Monthly Minimum", anchor="e")
        self.overview_tree.heading("per_pay", text="Per Pay Period", anchor="e")
        self.overview_tree.heading("free", text="Debt Free Estimated Date", anchor="center")
        self.overview_tree.column("type", width=140, anchor="w")
        self.overview_tree.column("remaining", width=120, anchor="e")
        self.overview_tree.column("monthly", width=150, anchor="e")
        self.overview_tree.column("per_pay", width=130, anchor="e")
        self.overview_tree.column("free", width=190, anchor="center")
        self.overview_tree.pack(fill="both", expand=True)
        self.overview_tree.tag_configure("alt", background=COLORS["alt_row"])
        self.overview_tree.tag_configure("total", background=COLORS["blue"])
        self.overview_types = types
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(1, weight=1)
        frm.grid_rowconfigure(2, weight=1)
        self.refresh_overview()

    def refresh_overview(self):
        if not hasattr(self, "overview_tree"):
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
        for item in self.overview_debtor_tree.get_children():
            self.overview_debtor_tree.delete(item)
        for item in self.overview_tree.get_children():
            self.overview_tree.delete(item)
        pay_periods = summary["pay_periods"] or 12
        for i, row in enumerate(summary.get("by_debtor") or []):
            tag = "alt" if i % 2 else ""
            self.overview_debtor_tree.insert("", "end", values=(
                row["name"],
                money(row["remaining"]),
                money(row["annual_min"] / 12.0),
                money(row["annual_min"] / pay_periods),
                format_display_date(row.get("debt_free") or "✓"),
            ), tags=(tag,))
        self.overview_debtor_tree.insert("", "end", values=(
            "Total",
            money(summary["remaining_total"]),
            money(summary["monthly_min"]),
            money(summary["per_pay"]),
            format_display_date(summary.get("debt_free_overall") or "✓"),
        ), tags=("total",))
        for i, name in enumerate(OVERVIEW_TYPES):
            row = summary["by_type"][name]
            tag = "alt" if i % 2 else ""
            self.overview_tree.insert("", "end", values=(
                name,
                money(row["remaining"]),
                money(row["annual_min"] / 12.0),
                money(row["annual_min"] / pay_periods),
                format_display_date(row.get("debt_free") or "✓"),
            ), tags=(tag,))
        self.overview_tree.insert("", "end", values=(
            "Total",
            money(summary["remaining_total"]),
            money(summary["monthly_min"]),
            money(summary["per_pay"]),
            format_display_date(summary.get("debt_free_overall") or "✓"),
        ), tags=("total",))

    # ----- Debt management -----
    def refresh_debt_tree(self):
        selected_id = None
        sel = self.debt_tree.selection()
        if sel:
            try:
                selected_id = int(sel[0])
            except (TypeError, ValueError):
                selected_id = None
        for item in self.debt_tree.get_children():
            self.debt_tree.delete(item)
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
        paid = [debt for debt in self.debts if not _is_active_debt(debt)]
        paid.sort(key=lambda debt: int(debt.get("id") or 0))
        sorted_debts = ranked + paid
        for i, d in enumerate(sorted_debts):
            is_paid = not _is_active_debt(d)
            if is_paid:
                tag = "zero_alt" if i % 2 else "zero"
                rank_text = "—"
                est_text = "✓"
            else:
                tag = "alt" if i % 2 else ""
                rank_text = str(d.get("payoff_rank") or "")
                est_text = format_display_date(payoff_by_id.get(d.get("id"), "Never"))
            current = parse_amount(d.get("current"), allow_blank=True) or 0.0
            rate = parse_amount(d.get("rate"), allow_blank=True) or 0.0
            minimum = parse_amount(d.get("min_payment"), allow_blank=True) or 0.0
            due_anchor, _due_frequency = due_series(d)
            due_text = format_display_date(due_anchor) if due_anchor else "—"
            self.debt_tree.insert("", "end", iid=str(d["id"]), values=(
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
            ), tags=(tag,))
        if selected_id is not None and str(selected_id) in self.debt_tree.get_children():
            self.debt_tree.selection_set(str(selected_id))
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
        if debt is None or not _is_active_debt(debt):
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

    def _on_debt_drag_press(self, event):
        if self.debt_tree.identify_region(event.x, event.y) not in ("cell", "tree"):
            self._debt_drag = None
            return
        item = self.debt_tree.identify_row(event.y)
        debt = self._debt_by_iid(item) if item else None
        if debt is None or not _is_active_debt(debt):
            self._debt_drag = None
            return
        self._debt_drag = {"item": item, "moved": False}

    def _on_debt_drag_motion(self, event):
        drag = self._debt_drag
        if not drag:
            return
        target = self.debt_tree.identify_row(event.y)
        if not target or target == drag["item"]:
            return
        target_debt = self._debt_by_iid(target)
        if target_debt is None or not _is_active_debt(target_debt):
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
            if debt is None or not _is_active_debt(debt):
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
        if not self.debts:
            return
        win = tk.Toplevel(self.root)
        self._prepare_dialog(win)
        win.title("Update Current Balances")
        win.geometry("520x450")
        win.configure(bg=COLORS["bg"])
        entries = {}
        for i, d in enumerate(self.debts):
            ttk.Label(win, text=d["short_name"]).grid(row=i, column=0, sticky="w", padx=8, pady=3)
            var = tk.StringVar(value=str(d["current"]))
            ent = ttk.Entry(win, textvariable=var, width=16)
            ent.grid(row=i, column=1, padx=8, pady=3)
            entries[d["id"]] = var
            ent.bind("<FocusOut>", lambda _event: commit())
            ent.bind("<Return>", lambda _event: commit())

        def commit():
            changes = []
            invalid = None
            for debt in self.debts:
                text = entries[debt["id"]].get().strip()
                value = parse_amount(text, minimum=0, maximum=1e12, allow_blank=True)
                if value is None:
                    invalid = debt.get("short_name") or "that debt"
                    continue
                current = parse_amount(debt.get("current"), allow_blank=True) or 0.0
                if abs(value - current) > 0.0001:
                    changes.append((debt, value))
            if invalid:
                self.status_var.set(f"Enter a non-negative number for {invalid}")
            if not changes:
                return
            self.push_undo()
            for debt, value in changes:
                debt["current"] = value
                self._sync_paid_off_state(debt)
            self.refresh_debt_tree()
            self.recalculate()

        ttk.Label(win, text="A valid balance is applied when you leave the field.").grid(
            row=len(self.debts), column=0, columnspan=2, padx=8, pady=12, sticky="w"
        )
        win.protocol("WM_DELETE_WINDOW", lambda: (commit(), win.destroy()))

    # ----- Settings -----
    def _sync_settings_form(self):
        self._settings_lock = True
        try:
            self.freq_var.set(self.settings.get("frequency", "biweekly"))
            self.date_var.set(format_display_date(self.settings.get("first_pay_date", "")))
            if hasattr(self, "anchor_date_var"):
                self.anchor_date_var.set(format_display_date(self.settings.get("balance_as_of") or self.settings.get("first_pay_date", "")))
            if hasattr(self, "starting_var"):
                self.starting_var.set(str(self.settings.get("starting_additional", 0)))
            self.proj_var.set(str(self.settings.get("projection_periods", 120)))
            if hasattr(self, "salary_var"):
                self.salary_var.set(str(self.settings.get("yearly_net_salary", 0)))
            if hasattr(self, "spouse_var"):
                self.spouse_var.set(str(self.settings.get("spouse_salary", 0)))
            if hasattr(self, "spouse_freq_var"):
                freq = self.settings.get("spouse_income_frequency", "biweekly")
                self.spouse_freq_var.set(freq if freq in PAY_FREQUENCIES else "biweekly")
            if hasattr(self, "other_income_var"):
                self.other_income_var.set(str(self.settings.get("other_income", 0)))
            if hasattr(self, "other_freq_var"):
                freq = self.settings.get("other_income_frequency", "biweekly")
                self.other_freq_var.set(freq if freq in PAY_FREQUENCIES else "biweekly")
        finally:
            self._settings_lock = False

    def _commit_settings(self, _event=None):
        if getattr(self, "_settings_lock", False):
            return
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
        self._sync_settings_form()
        self.refresh_debt_tree()
        self.refresh_archives()
        self.recalculate()
        self.status_var.set(f"Undo ({len(self.undo_stack)} left)")

    def redo(self):
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
        frm = ttk.Frame(parent, padding=12)
        frm.pack(fill="both", expand=True)
        self._chart_lock = False
        self._debtor_box_choice = None
        self._type_box_choice = None
        self._chart_filter_choice = None
        self._chart_selected_ids = set()
        self._chart_selection_custom = False
        self._chart_model = None

        cards = ttk.Frame(frm)
        cards.pack(fill="x")
        for column in range(3):
            cards.grid_columnconfigure(column, weight=1, uniform="chart_cards")

        def _box(column, title):
            frame = ttk.LabelFrame(cards, text=title, style="Overview.TLabelframe", padding=10)
            frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0 if column == 2 else 6))
            return frame

        all_box = _box(0, "All Debt")
        ttk.Label(all_box, text="Every debt with a balance", style="OverviewCaption.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        debtor_box = _box(1, "By Debtor")
        self.chart_debtor_var = tk.StringVar()
        self.chart_debtor_combo = ttk.Combobox(debtor_box, textvariable=self.chart_debtor_var, state="readonly", width=24)
        self.chart_debtor_combo.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.chart_debtor_combo.bind("<<ComboboxSelected>>", self._on_chart_debtor_box)
        type_box = _box(2, "By Debt Type")
        self.chart_type_var = tk.StringVar()
        self.chart_type_combo = ttk.Combobox(type_box, textvariable=self.chart_type_var, state="readonly", width=24)
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

        controls = ttk.Frame(frm)
        controls.pack(fill="x", pady=(12, 6))
        ttk.Label(controls, text="Category", style="Header.TLabel").pack(side="left", padx=(0, 8))
        self.chart_filter_var = tk.StringVar(value="Debt Type")
        self.chart_filter_combo = ttk.Combobox(
            controls, textvariable=self.chart_filter_var, state="readonly", width=16,
            values=("Debt Type", "Debtor"),
        )
        self.chart_filter_combo.pack(side="left")
        self.chart_filter_combo.bind("<<ComboboxSelected>>", self._on_chart_filter)

        ttk.Style().configure(
            "ChartTitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["header_bg"],
            font=("Segoe UI", 12, "bold"),
            anchor="center",
        )
        self.chart_title_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.chart_title_var, style="ChartTitle.TLabel", anchor="center").pack(fill="x", pady=(8, 2))

        chart_row = ttk.Frame(frm)
        chart_row.pack(fill="both", expand=True)
        chart_row.grid_columnconfigure(1, weight=1)
        chart_row.grid_rowconfigure(0, weight=1)
        self.chart_y_canvas = tk.Canvas(chart_row, width=84, bg="white", highlightthickness=0)
        self.chart_y_canvas.grid(row=0, column=0, sticky="ns")
        self.payoff_canvas = tk.Canvas(
            chart_row, bg="white", highlightthickness=1, highlightbackground=COLORS["border"],
        )
        self.payoff_canvas.grid(row=0, column=1, sticky="nsew")
        self.payoff_canvas.bind("<Configure>", self._on_chart_resize)
        self.payoff_canvas.bind("<MouseWheel>", self._on_chart_wheel)
        self.chart_y_canvas.bind("<MouseWheel>", self._on_chart_wheel)
        self.chart_legend_frame = ttk.Frame(chart_row)
        self.chart_legend_frame.grid(row=0, column=2, sticky="ns", padx=(10, 0))
        self.chart_legend_frame.grid_remove()
        ttk.Label(self.chart_legend_frame, text="Click to show or hide", style="OverviewCaption.TLabel").pack(anchor="w", pady=(2, 4))
        key_buttons = ttk.Frame(self.chart_legend_frame)
        key_buttons.pack(fill="x", pady=(0, 6))
        ttk.Style().configure("ChartKey.TButton", font=("Segoe UI", 9), padding=(8, 3))
        ttk.Button(key_buttons, text="Select all", style="ChartKey.TButton", command=self._select_all_chart_debts).pack(side="left")
        ttk.Button(key_buttons, text="Deselect all", style="ChartKey.TButton", command=self._clear_chart_debts).pack(side="left", padx=(6, 0))
        legend_body = ttk.Frame(self.chart_legend_frame)
        legend_body.pack(fill="both", expand=True)
        self.chart_legend_canvas = tk.Canvas(legend_body, width=1, bg=COLORS["bg"], highlightthickness=0, cursor="hand2")
        self.chart_legend_canvas.pack(side="left", fill="both", expand=True)
        self.chart_legend_scroll = ttk.Scrollbar(legend_body, orient="vertical", command=self.chart_legend_canvas.yview)
        self.chart_legend_canvas.configure(yscrollcommand=self._on_legend_scroll)
        self.chart_legend_canvas.bind("<Button-1>", self._on_chart_legend_click)
        self.chart_legend_canvas.bind("<MouseWheel>", self._on_legend_wheel)
        self.chart_xscroll = ttk.Scrollbar(chart_row, orient="horizontal", command=self.payoff_canvas.xview)
        self.chart_xscroll.grid(row=1, column=1, sticky="ew")
        self.chart_xscroll.grid_remove()
        self._chart_scrollable = False
        self._chart_bar_shown = False
        self._chart_legend_shown = False
        self._chart_legend_scroll_shown = False

    def _fill_interest_box(self, frame, minimum_var, additional_var, savings_var, note_var):
        ttk.Label(frame, text="Interest with minimum payments only", style="OverviewCaption.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Label(frame, textvariable=minimum_var, style="OverviewHero.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Interest with additional payments", style="OverviewCaption.TLabel").grid(row=3, column=0, sticky="w")
        ttk.Label(frame, textvariable=additional_var, style="OverviewHero.TLabel").grid(row=4, column=0, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Total savings", style="OverviewCaption.TLabel").grid(row=5, column=0, sticky="w")
        ttk.Label(frame, textvariable=savings_var, style="OverviewHero.TLabel").grid(row=6, column=0, sticky="w")
        ttk.Label(frame, textvariable=note_var, style="OverviewCaption.TLabel", wraplength=280).grid(row=7, column=0, sticky="w", pady=(6, 0))

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

    def _on_legend_scroll(self, first, last):
        if self._chart_legend_scroll_shown:
            self.chart_legend_scroll.set(first, last)

    def _on_legend_wheel(self, event):
        legend = self.chart_legend_canvas
        bbox = legend.bbox("all")
        if bbox and bbox[3] > legend.winfo_height() + 2:
            legend.yview_scroll(-1 if event.delta > 0 else 1, "units")
            return "break"
        return self._on_chart_wheel(event)

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

    def _on_chart_wheel(self, event):
        if not getattr(self, "_chart_scrollable", False):
            return
        direction = -1 if event.delta > 0 else 1
        self.payoff_canvas.xview_scroll(direction * 4, "units")
        return "break"

    def _set_chart_scroll(self, scroll, content_width, height, reset_scroll):
        canvas = self.payoff_canvas
        bar = self.chart_xscroll
        view = 0.0
        if scroll and not reset_scroll:
            try:
                view = canvas.xview()[0]
            except tk.TclError:
                view = 0.0
        canvas.configure(scrollregion=(0, 0, max(content_width, 1), max(height, 1)))
        if scroll:
            canvas.configure(xscrollcommand=bar.set)
            if not self._chart_bar_shown:
                bar.grid(row=1, column=1, sticky="ew")
                self._chart_bar_shown = True
            canvas.xview_moveto(0 if reset_scroll else view)
        else:
            canvas.configure(xscrollcommand="")
            canvas.xview_moveto(0)
            if self._chart_bar_shown:
                bar.grid_remove()
                self._chart_bar_shown = False
        self._chart_scrollable = scroll

    def _set_chart_legend(self, show, width):
        legend = self.chart_legend_canvas
        frame = self.chart_legend_frame
        if show:
            if int(float(legend.cget("width"))) != int(width):
                legend.configure(width=width)
            if not self._chart_legend_shown:
                frame.grid(row=0, column=2, sticky="ns", padx=(10, 0))
                self._chart_legend_shown = True
        elif self._chart_legend_shown:
            frame.grid_remove()
            self._chart_legend_shown = False
            if self._chart_legend_scroll_shown:
                self.chart_legend_scroll.pack_forget()
                self._chart_legend_scroll_shown = False

    def _sync_legend_scroll(self):
        legend = self.chart_legend_canvas
        bar = self.chart_legend_scroll
        bbox = legend.bbox("all")
        height = legend.winfo_height()
        if bbox:
            legend.configure(scrollregion=(0, 0, max(bbox[2], 1), bbox[3] + 6))
        overflow = bool(bbox and height > 20 and bbox[3] > height + 2)
        if overflow:
            if not self._chart_legend_scroll_shown:
                bar.pack(side="right", fill="y")
                self._chart_legend_scroll_shown = True
        else:
            legend.yview_moveto(0)
            if self._chart_legend_scroll_shown:
                bar.pack_forget()
                self._chart_legend_scroll_shown = False

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
                dataset = "All Debtors"
            else:
                series = chart_type_lines(snow["points"], self.debts)
                dataset = "All Debt Types"
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
                "dataset": dataset,
                "start": snow["start"],
                "frequency": snow["frequency"],
                "series": series,
            }
        finally:
            self._chart_lock = False
        self._draw_payoff_chart(reset_scroll=True)

    def _draw_payoff_chart(self, reset_scroll=False):
        canvas = getattr(self, "payoff_canvas", None)
        y_canvas = getattr(self, "chart_y_canvas", None)
        legend = getattr(self, "chart_legend_canvas", None)
        if canvas is None or y_canvas is None or legend is None:
            return
        self._chart_drawing = True
        try:
            canvas.delete("all")
            y_canvas.delete("all")
            legend.delete("all")
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            if width < 40 or height < 40:
                return
            model = getattr(self, "_chart_model", None) or {}
            series = model.get("series") or []
            if not series:
                self.chart_title_var.set("")
                self._set_chart_legend(False, 1)
                self._set_chart_scroll(False, width, height, True)
                canvas.create_text(
                    width / 2, height / 2, text="No debts with a balance to chart.",
                    fill=COLORS["text"], font=("Segoe UI", 11),
                )
                return
            selected = set(self._chart_selected_ids)
            visible = [item for item in series if item["id"] in selected]
            dataset = model.get("dataset") or ""
            self.chart_title_var.set(visible[0]["name"] if len(visible) == 1 else dataset)
            legend_width = 28 + max(110, min(200, max(len(item["name"]) for item in series) * 7))
            self._set_chart_legend(True, legend_width)
            row_h = 28
            for index, item in enumerate(series):
                tag = f"debt-{item['id']}"
                row_y = 4 + index * row_h
                turned_on = item["id"] in selected
                legend.create_rectangle(0, row_y, legend_width, row_y + row_h - 2, fill=COLORS["bg"], outline="", tags=(tag,))
                legend.create_oval(
                    4, row_y + 6, 18, row_y + 20,
                    fill=item["color"] if turned_on else COLORS["bg"],
                    outline=item["color"] if turned_on else "#b7c3ce",
                    width=2, tags=(tag,),
                )
                legend.create_text(
                    26, row_y + 13, text=item["name"], anchor="w", tags=(tag,),
                    fill=COLORS["text"] if turned_on else "#8aa0b2",
                    font=("Segoe UI", 10, "bold" if turned_on else "normal"),
                )
            self._sync_legend_scroll()
            start = model.get("start")
            if start is None:
                return
            if not visible:
                self._set_chart_scroll(False, width, height, True)
                canvas.create_text(
                    width / 2, height / 2, text="Click a debt in the key to show it.",
                    fill=COLORS["text"], font=("Segoe UI", 11),
                )
                return
            pad_left = 8
            pad_right = 24
            inner = max(width - pad_left - pad_right, 40)
            end, bulk_days, scroll, plot_width = chart_time_scale(start, visible, inner)
            model["end"] = end
            model["bulk_days"] = bulk_days
            model["scroll"] = scroll
            plot_left = pad_left
            plot_right = plot_left + plot_width
            content_width = plot_right + pad_right
            detailed = bulk_days < 365
            _used, x_ticks = chart_axis_ticks(
                start, end, model.get("frequency") or "biweekly", plot_width, detailed=detailed,
            )
            angled = (not detailed) or len(x_ticks) > 8
            plot_top = 16
            plot_bottom = max(plot_top + 40, height - (88 if angled else 48))
            plot_height = plot_bottom - plot_top
            self._set_chart_scroll(scroll, content_width if scroll else width, height, reset_scroll)
            y_max = 0.0
            for item in visible:
                for _day, amount in item["points"]:
                    if amount > y_max:
                        y_max = amount
            if y_max <= 0:
                y_max = 1.0
            magnitude = 10 ** math.floor(math.log10(y_max)) if y_max >= 1 else 1
            fraction = y_max / magnitude
            nice = 1 if fraction <= 1 else 2 if fraction <= 2 else 5 if fraction <= 5 else 10
            y_top = nice * magnitude
            if y_top < y_max:
                y_top = y_max
            span_days = max((end - start).days, 1)

            def x_of(day):
                if day < start:
                    day = start
                if day > end:
                    day = end
                return plot_left + (day - start).days / span_days * plot_width

            def y_of(amount):
                return plot_bottom - (max(0.0, amount) / y_top) * plot_height

            grid = "#e6eef5"
            canvas.create_rectangle(plot_left, plot_top, plot_right, plot_bottom, outline="", fill="#fbfcfe")
            y_canvas.create_line(83, plot_top, 83, plot_bottom, fill=COLORS["header_bg"])
            for step in range(5):
                value = y_top * step / 4
                y = y_of(value)
                canvas.create_line(plot_left, y, plot_right, y, fill=grid)
                label = f"${value:,.0f}" if value >= 100 else f"${value:,.2f}"
                y_canvas.create_text(74, y, text=label, anchor="e", fill="#243140", font=("Segoe UI", 9))
            for position, label in x_ticks:
                x = x_of(position)
                canvas.create_line(x, plot_top, x, plot_bottom, fill=grid)
            drawn = []
            for item in visible:
                color = item["color"]
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
                    fill=chart_mix_white(color, 0.88), outline="",
                )
            for color, coords in drawn:
                if len(coords) >= 4:
                    canvas.create_line(*coords, fill="#ffffff", width=5, capstyle="round", joinstyle="round")
                    canvas.create_line(*coords, fill=color, width=2, capstyle="round", joinstyle="round")
                canvas.create_oval(coords[0] - 3, coords[1] - 3, coords[0] + 3, coords[1] + 3, fill=color, outline="#ffffff", width=1)
                canvas.create_oval(coords[-2] - 4.5, coords[-1] - 4.5, coords[-2] + 4.5, coords[-1] + 4.5, fill="#ffffff", outline=color, width=2)
            canvas.create_rectangle(plot_left, plot_top, plot_right, plot_bottom, outline="#d5dee7")
            canvas.create_line(plot_left, plot_bottom, plot_right, plot_bottom, fill=COLORS["header_bg"])
            canvas.create_line(plot_left, plot_top, plot_left, plot_bottom, fill=COLORS["header_bg"])
            for position, label in x_ticks:
                x = x_of(position)
                canvas.create_line(x, plot_bottom, x, plot_bottom + 5, fill=COLORS["header_bg"])
                if angled:
                    canvas.create_text(
                        x, plot_bottom + 8, text=label, angle=32, anchor="ne",
                        fill="#243140", font=("Segoe UI", 9),
                    )
                else:
                    canvas.create_text(
                        x, plot_bottom + 8, text=label, anchor="n",
                        fill="#243140", font=("Segoe UI", 9),
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
            self.refresh_schedule_tree()
            self._refresh_snowball()
            self.refresh_debt_tree()
            self.refresh_overview()
            self.refresh_charts()
            return
        self.refresh_schedule_tree()
        self._refresh_snowball()
        self.refresh_debt_tree()
        self.refresh_overview()
        self.refresh_charts()
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

    def _render_period_grid(self, inner, periods, highlight_current=False, editable=False, include_hidden=False, debts=None, show_year=True, total_debt_last=False, frozen_inner=None, row_surface="#ffffff"):
        for widget in inner.winfo_children():
            widget.destroy()
        if frozen_inner is not None:
            for widget in frozen_inner.winfo_children():
                widget.destroy()
        if not periods:
            ttk.Label(inner, text="No pay periods in this year.").grid(row=0, column=0, sticky="w", padx=8, pady=8)
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

        def _header(parent, col_index, spec):
            lbl = tk.Label(
                parent, text=spec["title"], bg=COLORS["header_bg"], fg=COLORS["header_fg"],
                font=("Segoe UI", 9, "bold"), padx=3, pady=4, borderwidth=1, relief="solid",
                anchor=spec["header_anchor"],
            )
            lbl.grid(row=0, column=col_index, sticky="nsew")

        for index, spec in enumerate(frozen_cols):
            _header(frozen_inner, index, spec)
        for index, spec in enumerate(body_cols):
            _header(inner, index, spec)

        active_idx = None
        if highlight_current:
            active_pay = current_pay_date(self.settings)
            for i, per in enumerate(periods):
                try:
                    pay_date = parse_flexible_date(per.get("date"))
                except (TypeError, ValueError):
                    pay_date = None
                if pay_date == active_pay:
                    active_idx = per.get("index", i)
                    break

        def _money_cell(value):
            amount = parse_amount(value, allow_blank=True)
            if amount is None:
                return ""
            return f"${amount:,.2f}"

        def _plain_value(spec, period):
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
                return _money_cell(period.get("total_extra"))
            if key == "total":
                return _money_cell(period.get("total_debt"))
            return ""

        def _put_cell(parent, row, col, text, bg, fg, font, anchor, command=None):
            lbl = tk.Label(
                parent, text=text, bg=bg, fg=fg, font=font, padx=4, pady=2,
                borderwidth=0, relief="flat", highlightthickness=0, anchor=anchor,
            )
            lbl.grid(row=row, column=col, sticky="nsew")
            if command:
                lbl.bind("<Double-1>", command)
            return lbl

        for row, period in enumerate(periods, start=1):
            is_current = highlight_current and period.get("index") == active_idx
            row_bg = "#e3f2fd" if is_current else (row_surface if row % 2 == 0 else COLORS["alt_row"])
            font_weight = "bold" if is_current else "normal"
            base_font = ("Segoe UI", 9, font_weight)
            period_index = period.get("index")
            for col, spec in enumerate(frozen_cols):
                _put_cell(
                    frozen_inner, row, col, _plain_value(spec, period), row_bg, COLORS["text"],
                    base_font, spec["anchor"],
                    (lambda e, idx=period_index: self.edit_extras_for_period(idx)) if editable and period_index is not None else None,
                )
            for col, spec in enumerate(body_cols):
                if spec["key"] == "debt":
                    debt = spec["debt"]
                    raw = mapping_get(period.get("balances"), debt["id"], None)
                    if raw is None:
                        text, cell_bg, cell_fg, anchor, paid_font = "", row_bg, COLORS["text"], "center", base_font
                    else:
                        bal = parse_amount(raw, allow_blank=True)
                        if bal is None:
                            text, cell_bg, cell_fg, anchor, paid_font = "", row_bg, COLORS["text"], "center", base_font
                        elif bal < 0.5:
                            text, cell_bg, cell_fg, anchor, paid_font = "✓", COLORS["green"], "#146c43", "e", ("Segoe UI", 12, "bold")
                        else:
                            text, cell_bg, cell_fg, anchor, paid_font = f"${bal:,.2f}", row_bg, COLORS["text"], "e", base_font
                    command = None
                    if editable:
                        command = lambda e, idx=period["index"], did=debt["id"]: self.edit_extras_for_period(idx, focus_debt_id=did)
                    _put_cell(inner, row, col, text, cell_bg, cell_fg, paid_font, anchor, command)
                else:
                    _put_cell(
                        inner, row, col, _plain_value(spec, period), row_bg, COLORS["text"],
                        base_font, spec["anchor"],
                        (lambda e, idx=period["index"]: self.edit_extras_for_period(idx)) if editable else None,
                    )

        for col in range(len(frozen_cols)):
            frozen_inner.grid_columnconfigure(col, minsize=90)
            frozen_inner.grid_rowconfigure(0, minsize=28)
        for col in range(len(body_cols)):
            inner.grid_columnconfigure(col, minsize=90)
            inner.grid_rowconfigure(0, minsize=28)
        for row in range(1, len(periods) + 1):
            inner.grid_rowconfigure(row, minsize=24)
            if frozen_inner is not None:
                frozen_inner.grid_rowconfigure(row, minsize=24)

    def refresh_schedule_tree(self):
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
        fit = getattr(self, "_fit_schedule_panes", None)
        if fit:
            fit()

    def edit_extras_for_period(self, idx, focus_debt_id=None):
        if idx >= len(self.schedule):
            return
        period = self.schedule[idx]
        win = tk.Toplevel(self.root)
        self._prepare_dialog(win)
        title_extra = ""
        if focus_debt_id:
            dname = next((d["short_name"] for d in self.debts if d["id"] == focus_debt_id), "")
            title_extra = f" — {dname}"
        win.title(f"Extras for Period {period['period']} ({format_display_date(period['date'])}){title_extra}")
        win.geometry("480x520")
        win.configure(bg=COLORS["bg"])

        ttk.Label(win, text=f"Assign extra payments for {format_display_date(period['date'])}", style="Header.TLabel").pack(pady=8)
        entries = {}
        entry_widgets = {}
        closing = False

        bottom = ttk.Frame(win)
        bottom.pack(side="bottom", fill="x", pady=(4, 12))
        ttk.Label(bottom, text="A valid extra is applied when you leave the field or press Done.").pack(pady=(0, 8))
        done_btn = ttk.Button(bottom, text="Done", command=lambda: close_dialog())
        done_btn.pack()
        done_btn.bind("<ButtonPress-1>", lambda _event: remember_focused_entry(), add="+")

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12)

        visible_debts = self._schedule_column_debts()
        if not visible_debts:
            ttk.Label(frame, text="No debts on the pay schedule.").grid(row=0, column=0, sticky="w", pady=3)

        for i, d in enumerate(visible_debts):
            did = d["id"]
            period_map = mapping_get(self.extras_by_period, period["date"], {})
            if not isinstance(period_map, dict):
                period_map = {}
            current_extra = mapping_get(period_map, did, 0.0) or 0.0
            label_text = d["short_name"]
            if focus_debt_id and did == focus_debt_id:
                label_text = f"★ {label_text}"
            ttk.Label(frame, text=label_text).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(current_extra))
            ent = ttk.Entry(frame, textvariable=var, width=14)
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
            ttk.Label(self.archive_year_bar, text="No archived years yet.").pack(side="left", padx=2, pady=4)
            if hasattr(self, "archive_inner"):
                self._render_period_grid(self.archive_inner, [], row_surface=COLORS["bg"])
            return
        if selected is None or not any(int(year) == int(selected) for year in years):
            selected = int(years[0])
            self._archive_year = selected
        for year in years:
            year_number = int(year)
            chosen = year_number == int(selected)
            button = tk.Button(
                self.archive_year_bar,
                text=str(year_number),
                font=("Segoe UI", 10, "bold" if chosen else "normal"),
                bg=COLORS["header_bg"] if chosen else "white",
                fg=COLORS["header_fg"] if chosen else COLORS["header_bg"],
                activebackground=COLORS["header_bg"],
                activeforeground=COLORS["header_fg"],
                relief="solid",
                bd=1,
                padx=12,
                pady=3,
                cursor="hand2",
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
            row_surface=COLORS["bg"],
        )

    # ----- File I/O -----
    def new_blank(self):
        if messagebox.askyesno("New Blank", "Clear all debts and start blank?"):
            self._start_blank_plan()

    def _start_blank_plan(self):
        self.debts = []
        self.settings = deepcopy(DEFAULT_SETTINGS)
        today = datetime.date.today().isoformat()
        self.settings["first_pay_date"] = today
        self.settings["balance_as_of"] = today
        self.settings.pop("earliest_pay_date", None)
        self.extras_by_period = {}
        self.archives = {}
        self.current_file = None
        self._forget_remembered_file()
        self._set_title()
        self._sync_settings_form()
        self.refresh_debt_tree()
        self.recalculate()
        self.refresh_archives()
        self._mark_saved_state()
        self.status_var.set("Started blank")

    def _startup_prompts(self):
        if not self.current_file:
            self._prompt_missing_plan()
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        if not self._tutorial_was_seen():
            self.start_tutorial(auto=True)

    def _prompt_missing_plan(self):
        while True:
            choice = self._plan_choice_dialog()
            if choice == "open":
                if self.open_file():
                    return
                continue
            if choice == "new":
                self._start_blank_plan()
                return
            self.root.destroy()
            return

    def _plan_choice_dialog(self):
        choice = {"value": None}
        win = tk.Toplevel(self.root)
        self._prepare_dialog(win)
        win.title("Open or Create a Plan")
        win.geometry("460x180")
        win.configure(bg=COLORS["bg"])
        win.transient(self.root)
        ttk.Label(
            win,
            text="Create a new blank plan or open an existing JSON file. Those are the only choices.",
            wraplength=420,
            justify="left",
        ).pack(padx=20, pady=(20, 16), anchor="w")
        buttons = ttk.Frame(win)
        buttons.pack(pady=8)

        def choose_new():
            choice["value"] = "new"
            win.destroy()

        def choose_open():
            choice["value"] = "open"
            win.destroy()

        def choose_close():
            choice["value"] = None
            win.destroy()

        ttk.Button(buttons, text="Create New", command=choose_new).pack(side="left", padx=6)
        ttk.Button(buttons, text="Open Existing...", command=choose_open).pack(side="left", padx=6)
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
        return os.path.dirname(os.path.abspath(__file__))

    def _tutorial_seen_path(self):
        return os.path.join(self._user_data_dir(), "tutorial_seen.txt")

    def _tutorial_was_seen(self):
        return os.path.isfile(self._tutorial_seen_path())

    def _mark_tutorial_seen(self):
        try:
            with open(self._tutorial_seen_path(), "w", encoding="utf-8") as handle:
                handle.write("1")
        except OSError:
            pass

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
        if self.current_file:
            self.root.title(f"Debt Payoff Planner — {os.path.basename(self.current_file)}")
        else:
            self.root.title("Debt Payoff Planner")

    def _load_saved_plan(self, path):
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
        self.extras_by_period = sanitize_extras(data.get("extras_by_period", {}))
        self.archives = normalize_archives(data.get("archives", {}))
        self._freeze_archive_columns()
        self.current_file = path
        self._remember_file(path)
        self._set_title()
        self._sync_settings_form()
        self._hide_paid_off_from_schedule()
        self.refresh_debt_tree()
        self.recalculate()
        self.refresh_archives()
        self._mark_saved_state()
        self.status_var.set(f"Opened {os.path.basename(path)}")

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
        self.current_file = path
        if not self._write_json(path):
            return False
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
        for index in range(self.notebook.index("end")):
            if self.notebook.tab(index, "text") == name:
                self.notebook.select(index)
                return

    def _tutorial_steps(self):
        return [
            {
                "tab": "Overview",
                "target": "notebook",
                "title": "Welcome",
                "body": "This planner tracks debts, paydays, extras, and payoff dates. Work left to right: Settings for your paycheck and income, Debts for the loans, Pay Schedule for each payday, Overview for totals, Archives for history. File → Save writes a JSON plan. The last file you saved or opened comes back when you start the app.",
            },
            {
                "tab": "Settings",
                "target": "settings_freq",
                "title": "Pay frequency",
                "body": "Set how often you get paid: biweekly, semi-monthly, or monthly. Interest and the schedule step on this cadence. Leave a finished field (or press Enter) to apply it. Invalid text is left unchanged. There is no Apply button.",
            },
            {
                "tab": "Settings",
                "target": "settings_balance_as_of",
                "title": "The two dates",
                "body": "Balances As Of is the payday your Debts-tab current balances belong to. First Pay Date is only how far back the grid starts. Changing First Pay Date does not move the balances. The program may replace First Pay Date with the first payday of that year so period numbers line up.",
            },
            {
                "tab": "Settings",
                "target": "settings_starting",
                "title": "Snowball and projection",
                "body": "Starting Additional Per Pay is extra you can put toward debt each paycheck, before included paid-off payments are added. Uncheck a debt to leave its regular payment out. Projection Periods is how many future paydays the schedule shows.",
            },
            {
                "tab": "Settings",
                "target": "settings_salary",
                "title": "Income",
                "body": "Yearly Net Salary, Spouse / Secondary Yearly Net Income, and Other Yearly Net Salary are yearly net amounts. The spouse and other dropdowns only calculate Overview Pay Period Average from that yearly figure. Monthly Average is yearly net divided by 12. Totals and DTI use the yearly figures. Use 0 if a line does not apply.",
            },
            {
                "tab": "Debts",
                "target": "debt_tree",
                "title": "The debt list",
                "body": "Each row is one debt. Current Balance, Rate, and Minimum Payment are the live numbers the schedule uses. Short Name is the column title on the pay schedule. Green rows are paid off. They stay here until you delete them.",
            },
            {
                "tab": "Debts",
                "target": "debt_btn_frame",
                "title": "Adding and editing",
                "body": "Add Debt or Edit Debt opens the form. Choose a Debt Type. Next Due is MM-DD; Payment Frequency is how often that minimum repeats. Update Current Balances when a statement posts. A new debt starts on the current payday. Earlier schedule cells stay empty for that debt.",
            },
            {
                "tab": "Debts",
                "target": "debt_btn_frame",
                "title": "Payoff Rank",
                "body": "Rank 1 gets extra first, then 2, and so on. Drag an unpaid row, or select it and use Move Up / Move Down. When a balance is paid off that debt leaves the ranking, moves to the bottom, and the rest compact to 1, 2, 3. It is left out of remaining-debt estimates and the pay schedule. Its regular payment feeds Additional Per Pay unless you uncheck that option on the debt. A new debt, or a paid-off debt you reopen, goes last.",
            },
            {
                "tab": "Pay Schedule",
                "target": "sched_canvas",
                "title": "The pay grid",
                "body": "Each row is one payday. Period numbers restart each calendar year. Money columns are right-aligned with their titles. A green check is paid off. A blank cell means that debt had not started yet. Paid-off debts drop off this grid when you archive and when you open the plan.",
            },
            {
                "tab": "Pay Schedule",
                "target": "sched_btn_frame",
                "title": "Extras",
                "body": "Double-click a cell to assign extra for that payday. Only debts still on this tab are listed. Type the amount, then press Done so the field you were editing is saved. These extras are manual. They stay on that pay date.",
            },
            {
                "tab": "Pay Schedule",
                "target": "_snowball",
                "title": "Additional Per Pay",
                "body": "This box is the snowball recommendation: Starting Additional Per Pay plus regular payments from debts already paid off, spread across your paychecks. A debt can be unchecked so its payment is left out. It is not applied automatically on the grid. Overview uses it on paydays that have not arrived. A payday that has started uses the extra logged on that row. If that extra is missing, the payoff date moves later. If more than expected is logged, the date moves sooner.",
            },
            {
                "tab": "Pay Schedule",
                "target": "sched_archive_btn",
                "title": "Archive",
                "body": "Archive Completed Pay Periods moves past paydays into Archives. Remaining balances become the new current amounts. That cannot be undone except with Undo, and a later edit does not put those rows back on the live schedule.",
            },
            {
                "tab": "Overview",
                "target": "overview_income",
                "title": "Income and DTI",
                "body": "Income shows the yearly amounts and their monthly averages. Front-end DTI is mortgage minimums only (N/A with no mortgage). Back-end DTI is all active monthly minimums. Both need income; otherwise they show a dash. Paid-off debts are left out.",
            },
            {
                "tab": "Overview",
                "target": "overview_types",
                "title": "By debtor, by type, and payoff dates",
                "body": "By Debtor lists each debt. Debt by Type rolls those same debts up by type. Remaining balance and minimums match, and both estimated dates match the Debts tab. Changing a balance, or logging a larger, smaller, or missing extra on the Pay Schedule, moves the dates in both boxes.",
            },
            {
                "tab": "Archives",
                "target": "archive_year_bar",
                "title": "Archives",
                "body": "Click a year to see the same columns as the pay schedule. Archived periods never change, even if you delete the debt later. That is the history of what was on the schedule at archive time.",
            },
        ]

    def start_tutorial(self, auto=False):
        self._tutorial_auto = bool(auto)
        win = getattr(self, "_tutorial_win", None)
        if win is not None:
            try:
                if win.winfo_exists():
                    self._tutorial_index = 0
                    self._show_tutorial_step()
                    return
            except tk.TclError:
                pass
        self._tutorial_index = 0
        self._build_tutorial_window()
        self._show_tutorial_step()

    def _build_tutorial_window(self):
        win = tk.Toplevel(self.root)
        self._tutorial_win = win
        self._prepare_dialog(win)
        win.overrideredirect(True)
        win.configure(bg=COLORS["header_bg"])
        win.protocol("WM_DELETE_WINDOW", self._close_tutorial)

        shell = tk.Frame(win, bg=COLORS["header_bg"])
        shell.pack(fill="both", expand=True)

        self._tutorial_pointer = tk.Canvas(shell, width=16, height=16, bg=COLORS["header_bg"], highlightthickness=0)
        self._tutorial_body = tk.Frame(shell, bg="white", padx=14, pady=12)
        header = tk.Frame(self._tutorial_body, bg="white")
        header.pack(fill="x", pady=(0, 6))
        self._tutorial_title = tk.Label(
            header, text="", bg="white", fg=COLORS["header_bg"],
            font=("Segoe UI", 12, "bold"), anchor="w", justify="left", wraplength=310,
        )
        self._tutorial_title.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._tutorial_mascot = tk.Canvas(header, width=72, height=72, bg="white", highlightthickness=0)
        self._tutorial_mascot.pack(side="right")
        self._tutorial_text = tk.Label(
            self._tutorial_body, text="", bg="white", fg=COLORS["text"],
            font=("Segoe UI", 10), wraplength=380, justify="left", anchor="nw",
        )
        self._tutorial_text.pack(fill="both", expand=True)
        footer = tk.Frame(self._tutorial_body, bg="white")
        footer.pack(fill="x", pady=(12, 0))
        self._tutorial_progress = tk.Label(footer, text="", bg="white", fg="#5a6a7a", font=("Segoe UI", 9))
        self._tutorial_progress.pack(side="left")
        self._tutorial_skip = ttk.Button(footer, text="Skip", command=self._skip_tutorial)
        ttk.Button(footer, text="Close", command=self._close_tutorial).pack(side="right")
        self._tutorial_next = ttk.Button(footer, text="Next", command=self._tutorial_next_step)
        self._tutorial_next.pack(side="right", padx=(0, 6))
        self._tutorial_back = ttk.Button(footer, text="Back", command=self._tutorial_back_step)
        self._tutorial_back.pack(side="right", padx=(0, 6))

        self._tutorial_pack = {"shell": shell, "side": None}
        try:
            win.grab_set()
        except tk.TclError:
            pass

    def _skip_tutorial(self):
        self._close_tutorial()

    def _tutorial_next_step(self):
        steps = self._tutorial_steps()
        if self._tutorial_index >= len(steps) - 1:
            self._close_tutorial()
            return
        self._tutorial_index += 1
        self._show_tutorial_step()

    def _tutorial_back_step(self):
        if self._tutorial_index <= 0:
            return
        self._tutorial_index -= 1
        self._show_tutorial_step()

    def _close_tutorial(self):
        if getattr(self, "_tutorial_auto", False):
            self._mark_tutorial_seen()
        self._tutorial_auto = False
        self._stop_tutorial_mascot()
        win = getattr(self, "_tutorial_win", None)
        self._tutorial_win = None
        if win is not None:
            try:
                win.grab_release()
            except tk.TclError:
                pass
            try:
                win.destroy()
            except tk.TclError:
                pass
        self._align_snowball()

    def _show_tutorial_step(self):
        win = getattr(self, "_tutorial_win", None)
        if win is None or not win.winfo_exists():
            return
        steps = self._tutorial_steps()
        step = steps[self._tutorial_index]
        self._select_tab(step["tab"])
        self.root.update_idletasks()
        self._align_snowball()
        self._tutorial_title.config(text=step["title"])
        self._tutorial_text.config(text=step["body"])
        self._tutorial_progress.config(text=f"{self._tutorial_index + 1} of {len(steps)}")
        self._tutorial_back.state(["!disabled"] if self._tutorial_index else ["disabled"])
        self._tutorial_next.config(text="Done" if self._tutorial_index == len(steps) - 1 else "Next")
        if self._tutorial_index == 0:
            self._tutorial_skip.pack(side="left", padx=(12, 0))
        else:
            self._tutorial_skip.pack_forget()
        self._start_tutorial_mascot(step["tab"])
        target = getattr(self, step["target"], None)
        self.root.after(50, lambda: self._place_tutorial_bubble(target))

    def _place_tutorial_bubble(self, target):
        win = getattr(self, "_tutorial_win", None)
        if win is None:
            return
        try:
            if not win.winfo_exists():
                return
        except tk.TclError:
            return
        win.update_idletasks()
        body_w = max(self._tutorial_body.winfo_reqwidth(), 420)
        body_h = max(self._tutorial_body.winfo_reqheight(), 200)
        arrow = 16
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        side = "left"
        x = self.root.winfo_rootx() + 40
        y = self.root.winfo_rooty() + 80
        aim_y = y + body_h // 2
        try:
            if target is not None and target.winfo_exists() and target.winfo_ismapped():
                tx, ty = target.winfo_rootx(), target.winfo_rooty()
                tw, th = max(target.winfo_width(), 1), max(target.winfo_height(), 1)
                aim_y = ty + th // 2
                if tx + tw + arrow + body_w + 8 <= screen_w:
                    side = "left"
                    x = tx + tw
                    y = aim_y - body_h // 2
                elif tx - arrow - body_w - 8 >= 0:
                    side = "right"
                    x = tx - body_w - arrow
                    y = aim_y - body_h // 2
                elif ty + th + arrow + body_h + 8 <= screen_h:
                    side = "top"
                    x = tx + tw // 2 - body_w // 2
                    y = ty + th
                else:
                    side = "bottom"
                    x = tx + tw // 2 - body_w // 2
                    y = ty - body_h - arrow
        except tk.TclError:
            pass
        y = max(8, min(y, screen_h - body_h - arrow - 8))
        x = max(8, min(x, screen_w - body_w - arrow - 8))
        if side in ("left", "right"):
            along = aim_y - y
        else:
            try:
                along = int(target.winfo_rootx() + target.winfo_width() / 2 - x) if target is not None else body_w // 2
            except tk.TclError:
                along = body_w // 2
        self._layout_tutorial_pointer(side, body_w, body_h, along)
        total_w = body_w + (arrow if side in ("left", "right") else 0)
        total_h = body_h + (arrow if side in ("top", "bottom") else 0)
        win.geometry(f"{total_w}x{total_h}+{int(x)}+{int(y)}")
        win.lift()

    def _layout_tutorial_pointer(self, side, body_w, body_h, along):
        pointer = self._tutorial_pointer
        body = self._tutorial_body
        pointer.pack_forget()
        body.pack_forget()
        pointer.delete("all")
        arrow = 16
        fill = "white"
        along = max(18, min(int(along), (body_h if side in ("left", "right") else body_w) - 18))
        if side == "left":
            pointer.config(width=arrow, height=body_h)
            pointer.pack(side="left", fill="y")
            body.pack(side="left", fill="both", expand=True)
            pointer.create_polygon(0, along, arrow, along - 10, arrow, along + 10, fill=fill, outline=fill)
        elif side == "right":
            body.pack(side="left", fill="both", expand=True)
            pointer.config(width=arrow, height=body_h)
            pointer.pack(side="left", fill="y")
            pointer.create_polygon(arrow, along, 0, along - 10, 0, along + 10, fill=fill, outline=fill)
        elif side == "top":
            pointer.config(width=body_w, height=arrow)
            pointer.pack(side="top", fill="x")
            body.pack(side="top", fill="both", expand=True)
            pointer.create_polygon(along, 0, along - 10, arrow, along + 10, arrow, fill=fill, outline=fill)
        else:
            body.pack(side="top", fill="both", expand=True)
            pointer.config(width=body_w, height=arrow)
            pointer.pack(side="top", fill="x")
            pointer.create_polygon(along, arrow, along - 10, 0, along + 10, 0, fill=fill, outline=fill)
        self._tutorial_pack["side"] = side

    def _stop_tutorial_mascot(self):
        handle = getattr(self, "_tutorial_anim_id", None)
        if handle is not None:
            try:
                self.root.after_cancel(handle)
            except tk.TclError:
                pass
        self._tutorial_anim_id = None

    def _start_tutorial_mascot(self, tab):
        self._stop_tutorial_mascot()
        self._tutorial_anim_tab = tab
        self._tutorial_anim_frame = 0
        self._draw_tutorial_mascot()
        self._tutorial_anim_id = self.root.after(110, self._tick_tutorial_mascot)

    def _tick_tutorial_mascot(self):
        win = getattr(self, "_tutorial_win", None)
        if win is None:
            return
        try:
            if not win.winfo_exists():
                return
        except tk.TclError:
            return
        self._tutorial_anim_frame = (getattr(self, "_tutorial_anim_frame", 0) + 1) % 12
        self._draw_tutorial_mascot()
        self._tutorial_anim_id = self.root.after(110, self._tick_tutorial_mascot)

    def _draw_tutorial_mascot(self):
        canvas = getattr(self, "_tutorial_mascot", None)
        if canvas is None:
            return
        canvas.delete("all")
        tab = getattr(self, "_tutorial_anim_tab", "Overview")
        frame = getattr(self, "_tutorial_anim_frame", 0)
        if tab == "Settings":
            self._draw_mascot_gear_coin(canvas, frame)
        elif tab == "Debts":
            self._draw_mascot_snowball(canvas, frame)
        elif tab == "Pay Schedule":
            self._draw_mascot_paycheck(canvas, frame)
        elif tab == "Archives":
            self._draw_mascot_vault(canvas, frame)
        else:
            self._draw_mascot_chip(canvas, frame)

    def _draw_mascot_chip(self, canvas, frame):
        """Overview: a gold coin that bobs and blinks."""
        bob = (0, 1, 2, 1, 0, -1, -2, -1, 0, 1, 2, 1)[frame]
        cy = 36 + bob
        canvas.create_oval(10, cy - 26, 62, cy + 26, fill="#e8c547", outline="#b8860b", width=2)
        canvas.create_oval(16, cy - 20, 56, cy + 20, outline="#d4af37", width=1)
        canvas.create_text(36, cy + 1, text="$", fill=COLORS["header_bg"], font=("Segoe UI", 18, "bold"))
        blink = frame % 12 in (4, 5)
        if blink:
            canvas.create_line(26, cy - 8, 32, cy - 8, fill=COLORS["header_bg"], width=2)
            canvas.create_line(40, cy - 8, 46, cy - 8, fill=COLORS["header_bg"], width=2)
        else:
            canvas.create_oval(26, cy - 11, 32, cy - 5, fill=COLORS["header_bg"], outline="")
            canvas.create_oval(40, cy - 11, 46, cy - 5, fill=COLORS["header_bg"], outline="")
        canvas.create_arc(24, cy + 2, 48, cy + 16, start=200, extent=140, style="arc", outline=COLORS["header_bg"], width=2)

    def _draw_mascot_gear_coin(self, canvas, frame):
        """Settings: a coin with turning gear teeth."""
        cx, cy = 36, 38
        canvas.create_oval(18, 18, 54, 54, fill="#e8c547", outline="#b8860b", width=2)
        canvas.create_text(cx, cy + 1, text="$", fill=COLORS["header_bg"], font=("Segoe UI", 14, "bold"))
        teeth = 8
        turn = frame * 8
        for i in range(teeth):
            angle = math.radians(turn + i * (360 / teeth))
            x1 = cx + 22 * math.cos(angle)
            y1 = cy + 22 * math.sin(angle)
            x2 = cx + 30 * math.cos(angle)
            y2 = cy + 30 * math.sin(angle)
            canvas.create_line(x1, y1, x2, y2, fill="#7f8c9a", width=3, capstyle="round")
        wink = frame % 12 in (7, 8)
        if wink:
            canvas.create_line(28, 28, 32, 28, fill=COLORS["header_bg"], width=2)
        else:
            canvas.create_oval(27, 26, 33, 32, fill=COLORS["header_bg"], outline="")

    def _draw_mascot_snowball(self, canvas, frame):
        """Debts: a rolling snowball that grows a little as it turns."""
        shift = (0, 2, 4, 2, 0, -2, -4, -2, 0, 2, 4, 2)[frame]
        grow = 1 if frame % 6 < 3 else 0
        cx, cy = 36 + shift // 2, 40
        r = 22 + grow
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r, fill="#f7fbff", outline="#8aa4b8", width=2)
        canvas.create_oval(cx - 8 + shift // 3, cy - 10, cx - 2 + shift // 3, cy - 4, outline="#b9c9d4")
        canvas.create_oval(cx + 2, cy + 2, cx + 10, cy + 10, outline="#b9c9d4")
        canvas.create_oval(cx - 12, cy + 6, cx - 6, cy + 12, fill="#dbe7ef", outline="")
        canvas.create_oval(cx - 6, cy - 8, cx, cy - 2, fill=COLORS["header_bg"], outline="")
        canvas.create_oval(cx + 4, cy - 8, cx + 10, cy - 2, fill=COLORS["header_bg"], outline="")
        canvas.create_arc(cx - 10, cy - 2, cx + 10, cy + 12, start=200, extent=140, style="arc", outline=COLORS["header_bg"], width=2)
        canvas.create_line(8, 64, 64, 64, fill="#c5d0d8", width=2)

    def _draw_mascot_paycheck(self, canvas, frame):
        """Pay Schedule: an envelope paycheck that flutters."""
        lift = (0, 1, 2, 1, 0, -1, -2, -1, 0, 1, 2, 1)[frame]
        y = 18 + lift
        canvas.create_rectangle(10, y + 10, 62, y + 42, fill="#f4f7ee", outline=COLORS["header_bg"], width=2)
        canvas.create_polygon(10, y + 10, 36, y + 26, 62, y + 10, fill="#3DDC67", outline=COLORS["header_bg"], width=2)
        canvas.create_text(36, y + 32, text="PAY", fill=COLORS["header_bg"], font=("Segoe UI", 8, "bold"))
        if frame % 6 < 3:
            canvas.create_oval(52, y + 4, 64, y + 16, fill="#e8c547", outline="#b8860b")
            canvas.create_text(58, y + 10, text="$", fill=COLORS["header_bg"], font=("Segoe UI", 7, "bold"))

    def _draw_mascot_vault(self, canvas, frame):
        """Archives: a tiny vault whose wheel turns."""
        canvas.create_rectangle(14, 16, 58, 60, fill="#4a6280", outline=COLORS["header_bg"], width=2)
        canvas.create_rectangle(18, 20, 54, 56, fill="#2c3e50", outline="#1a365d")
        cx, cy = 36, 38
        canvas.create_oval(cx - 12, cy - 12, cx + 12, cy + 12, outline="#e8c547", width=2)
        turn = frame * 15
        for i in range(4):
            angle = math.radians(turn + i * 90)
            canvas.create_line(
                cx, cy,
                cx + 10 * math.cos(angle),
                cy + 10 * math.sin(angle),
                fill="#e8c547", width=2,
            )
        canvas.create_oval(cx - 3, cy - 3, cx + 3, cy + 3, fill="#e8c547", outline="")
        canvas.create_rectangle(48, 34, 58, 42, fill="#b8860b", outline=COLORS["header_bg"])

    def show_help(self):
        help_text = """Debt Payoff Planner - Quick Guide

1. Overview tab
   • Income shows yearly, pay-period average, and monthly average. Spouse / Secondary Yearly Net Income and Other Yearly Net Salary are yearly figures. Their Settings dropdowns only calculate Overview Pay Period Average from that yearly amount. Those averages are Overview only.
   • Front-end DTI is mortgage minimums only. N/A if there is no mortgage. A dash until income is entered
   • Back-end DTI is all active monthly minimums divided by average monthly income
   • Paid-off debts are left out. Debts without a type are Unspecified
   • Each debt is estimated in Payoff Rank order with its minimum plus accumulated extra. A type’s date is the last of those debts
   • When a debt is paid off in that estimate, its regular payment is added to Additional Per Pay on later paydays, unless that debt is set to leave it out

2. Pay Schedule tab
   • Running balance for each debt each period
   • Double-click a cell to assign an extra. Press Done to save the field you were editing
   • Green checks are paid-off balances. Empty cells are periods before a debt started
   • Additional Per Pay on the tab strip is starting additional plus regular payments from paid-off debts that are still included
   • Archive Completed Pay Periods moves past rows to Archives

3. Debts tab
   • Add / Edit / Delete, or Update Current Balances
   • Debt Type: Auto, Personal, Mortgage, Credit, or Payment Plan
   • Payment Frequency is how often that minimum repeats
   • Next Due is MM-DD. The year is the next time that date occurs
   • Payoff Rank is extra order. Drag a row, or Move Up / Move Down. Estimated Payoff is that snowball date. A new debt is last
   • A paid-off debt leaves the ranking, moves to the bottom, and is left out of remaining-debt estimates. Remaining unpaid debts are reranked 1, 2, 3
   • Green rows are paid off. They stay here until you delete them

4. Archives tab
   • Choose a year to see the same columns as the pay schedule
   • Archived periods do not change later, including if you delete the debt

5. Settings tab
   • Pay frequency: biweekly / semimonthly / monthly
   • Balances As Of is the payday your current balances belong to
   • First Pay Date is how far back the grid starts. It does not move the balances
   • Starting Additional Per Pay is the snowball base
   • Yearly income fields feed Overview
   • Leave a completed field to apply it. Letters, negatives, and bad dates are rejected and the previous value is kept

6. Charts tab
   • Three boxes show interest with minimum payments only, interest with additional payments in payoff order, and the interest saved
   • All Debt covers every balance. By Debtor opens on the next debt in payoff order. By Debt Type opens on the next type to reach zero
   • The Category dropdown switches the key between Debt Type and Debtor
   • Click a name to show or hide that line. Select all and Deselect all are above the key. The smallest balance starts on, and more than one can be on
   • A payoff of a year or more uses month names, such as December '26. Under a year, the labels are the plan's pay dates
   • If one shown debt runs much further than the rest, the chart scrolls so those closer dates stay readable
   • Paid-off debts and types are left out

7. Saving
   • File → Save / Save As writes JSON. The last file is reopened next time
   • File → New Blank starts clean
   • Undo / Redo cover debts, extras, settings, and archives

Interest is: balance × (annual rate ÷ pay periods per year)
"""
        messagebox.showinfo("How to Use", help_text)


class DebtDialog(tk.Toplevel):
    def __init__(self, parent, title="Debt", debt=None):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()
        prepare = getattr(parent, "_prepare_dialog", None)
        if prepare:
            prepare(self)
        self.configure(bg=COLORS["bg"])

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
        for key, label, default in fields_before_type:
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=4)
            val = str(debt.get(key, default)) if debt else default
            var = tk.StringVar(value=val)
            ttk.Entry(self, textvariable=var, width=32).grid(row=row, column=1, padx=10, pady=4)
            self.vars[key] = var
            row += 1

        ttk.Label(self, text="Debt Type").grid(row=row, column=0, sticky="w", padx=10, pady=4)
        current_type = (debt or {}).get("debt_type") or ""
        if current_type not in DEBT_TYPES:
            current_type = ""
        self.type_var = tk.StringVar(value=current_type)
        type_combo = ttk.Combobox(self, textvariable=self.type_var, values=DEBT_TYPES, state="readonly", width=29)
        type_combo.grid(row=row, column=1, padx=10, pady=4)
        row += 1

        for key, label, default in fields_after_type:
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=4)
            if debt and key == "due_md":
                val = debt.get("due_md") or ""
                if not val:
                    parsed = parse_month_day(debt.get("due_date"))
                    val = f"{parsed[0]:02d}-{parsed[1]:02d}" if parsed else ""
            else:
                val = str(debt.get(key, default)) if debt else default
            var = tk.StringVar(value=val)
            ttk.Entry(self, textvariable=var, width=32).grid(row=row, column=1, padx=10, pady=4)
            self.vars[key] = var
            row += 1

        freq_row = row
        ttk.Label(self, text="Payment Frequency").grid(row=freq_row, column=0, sticky="w", padx=10, pady=4)
        current_freq = (debt or {}).get("due_frequency") or "monthly"
        self.freq_var = tk.StringVar(value=DUE_FREQUENCY_LABELS.get(current_freq, "Monthly"))
        freq_combo = ttk.Combobox(
            self,
            textvariable=self.freq_var,
            values=[DUE_FREQUENCY_LABELS[key] for key in DUE_FREQUENCIES],
            state="readonly",
            width=29,
        )
        freq_combo.grid(row=freq_row, column=1, padx=10, pady=4)

        self.include_paid_var = tk.BooleanVar(value=include_paid_minimum(debt) if debt else True)
        ttk.Checkbutton(
            self,
            text="Include regular payment in Additional Per Pay when paid off",
            variable=self.include_paid_var,
        ).grid(row=freq_row + 1, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 2))

        btn = ttk.Frame(self)
        btn.grid(row=freq_row + 2, column=0, columnspan=2, pady=12)
        ttk.Button(btn, text="OK", command=self.on_ok).pack(side="left", padx=6)
        ttk.Button(btn, text="Cancel", command=self.destroy).pack(side="left", padx=6)

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


def main():
    root = tk.Tk()
    app = DebtPayoffApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
