import re
from collections import defaultdict
from copy import copy
from datetime import datetime, date, time, timedelta, timezone
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo
from urllib.parse import urlparse

import requests
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================
# APP SETTINGS
# =========================

BASE_URL = "https://results.advancedeventsystems.com/api/event"
MATCH_ENDPOINT_CONSTANT = "240"

TEMPLATE_PATH = "GRID EXAMPLE.xlsx"
TIMEZONE = "America/New_York"

ROLE_ROWS = ["match", "format", "R1", "R2", "LJ", "LJ", "SK", "AS"]

WEEKDAY_ABBREVIATIONS = {
    0: "MO",
    1: "TU",
    2: "WE",
    3: "TH",
    4: "FR",
    5: "SA",
    6: "SU",
}

DEFAULT_LEVEL_ABBREVIATIONS = {
    "OPEN": "O",
    "PREMIER": "P",
    "ELITE": "E",
    "SELECT": "S",
    "ASCEND": "A",
    "CLUB": "C",
    "ASPIRE": "AS",
    "SPIRIT": "SP",
    "CLASSIC": "CL",
    "POWER": "P",
    "P": "P",
    "GOLD": "G",
    "SILVER": "S",
    "BRONZE": "B",
    "PLATINUM": "P",
    "DIAMOND": "D",
    "RUBY": "R",
    "SAPPHIRE": "S",
    "EMERALD": "E",
    "NATIONAL": "N",
    "AMERICAN": "A",
    "REGIONAL": "R",
    "USA": "U",
    "GIRLS": "U",
    "GIRL": "U",
    "U": "U",
}


# =========================
# INPUT HELPERS
# =========================

def extract_event_key(user_input: str) -> str:
    text = (user_input or "").strip()

    if not text:
        return ""

    if text.startswith("http://") or text.startswith("https://"):
        parsed = urlparse(text)
        path_parts = [part for part in parsed.path.split("/") if part]

        if "event" in path_parts:
            event_index = path_parts.index("event")
            if event_index + 1 < len(path_parts):
                return path_parts[event_index + 1].strip()

    match = re.search(r"/event/([^/?#\s]+)", text)
    if match:
        return match.group(1).strip()

    match = re.search(r"/api/event/([^/?#\s]+)", text)
    if match:
        return match.group(1).strip()

    return text


def reset_app():
    keys_to_clear = [
        "event_info",
        "event_key",
        "generated_workbooks",
        "level_rows",
        "level_editor",
        "event_match_counts",
        "comparison_result",
        "comparison_uploaded_summary",
        "comparison_selected_date",
        "changed_court_workbooks",
    ]

    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]

    st.rerun()


# =========================
# API HELPERS
# =========================

def fetch_json(url: str) -> dict:
    headers = {
        "accept": "application/json, text/plain, */*",
        "user-agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    }

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.json()


def get_event_info(event_key: str) -> dict:
    return fetch_json(f"{BASE_URL}/{event_key}")


def get_match_info(event_key: str, event_date: date) -> dict:
    date_string = event_date.strftime("%Y-%m-%d")
    return fetch_json(
        f"{BASE_URL}/{event_key}/courts/{date_string}/{MATCH_ENDPOINT_CONSTANT}"
    )


# =========================
# DATE / TIME HELPERS
# =========================

def parse_event_date(value: str) -> date:
    return datetime.fromisoformat(value.split(".")[0]).date()


def date_range(start_date: date, end_date: date):
    current = start_date

    while current <= end_date:
        yield current
        current += timedelta(days=1)


def format_long_date(d: date) -> str:
    return f"{d.strftime('%A')}, {d.strftime('%B')} {d.day}"


def sheet_name_for_date(d: date) -> str:
    weekday = WEEKDAY_ABBREVIATIONS[d.weekday()]
    return f"{weekday} {d.month}-{d.day}"


def c1_label_for_date(d: date) -> str:
    weekday = WEEKDAY_ABBREVIATIONS[d.weekday()]
    return f"{weekday} {d.month}/{d.day}"


def ms_to_local_datetime(ms: int, tz_name: str) -> datetime:
    tz = ZoneInfo(tz_name)
    return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).astimezone(tz)


def timestamp_suffix(tz_name: str) -> str:
    now = datetime.now(ZoneInfo(tz_name))
    return now.strftime("__%y%m%d_%H%M")


def round_up_to_half_hour(dt: datetime) -> datetime:
    dt = dt.replace(second=0, microsecond=0)

    if dt.minute == 0 or dt.minute == 30:
        return dt

    if dt.minute < 30:
        return dt.replace(minute=30)

    return dt.replace(minute=0) + timedelta(hours=1)


def round_match_start_to_next_grid_time(match_start: datetime, grid_times: list[datetime]) -> datetime | None:
    clean_start = match_start.replace(second=0, microsecond=0)

    for grid_time in grid_times:
        if grid_time >= clean_start:
            return grid_time

    return None


def match_duration_minutes(match: dict) -> int:
    start_ms = match["ScheduledStartDateTime"]
    end_ms = match["ScheduledEndDateTime"]

    return round((end_ms - start_ms + 1) / 1000 / 60)


def determine_match_format(match: dict) -> str:
    minutes = match_duration_minutes(match)

    if minutes < 60:
        return "1 set to 25"

    if minutes == 60:
        return "best 2/3"

    return "auto-3"


# =========================
# DIVISION / MATCH NAME HELPERS
# =========================

def normalize_power_text(text: str) -> str:
    if not text:
        return ""

    text = re.sub(r"\bPower\b", "P", text, flags=re.IGNORECASE)
    text = re.sub(r"Power", "P", text, flags=re.IGNORECASE)

    return text


def normalize_level_name(level_name: str) -> str:
    level_name = normalize_power_text((level_name or "").strip())
    level_name = re.sub(r"\s+", " ", level_name)
    return level_name


def normalize_hex_color(hex_color: str | None) -> str:
    if not hex_color:
        return "FFFFFFFF"

    clean = str(hex_color).strip().replace("#", "").upper()

    if len(clean) == 6:
        return f"FF{clean}"

    if len(clean) == 8:
        return clean

    return "FFFFFFFF"


def display_hex_color(hex_color: str | None) -> str:
    if not hex_color:
        return "#FFFFFF"

    clean = str(hex_color).strip().replace("#", "").upper()

    if len(clean) == 6:
        return f"#{clean}"

    if len(clean) == 8:
        return f"#{clean[-6:]}"

    return "#FFFFFF"


def division_age(division_name: str) -> str:
    match = re.search(r"^\s*(\d{1,2})", division_name or "")
    return match.group(1) if match else ""


def division_level_name(division_name: str) -> str:
    name = normalize_power_text((division_name or "").strip())
    name = re.sub(r"^\d{1,2}\s*", "", name).strip()
    name = normalize_level_name(name)

    return name or "Division"


def default_level_abbreviation(level_name: str) -> str:
    level = normalize_level_name(level_name)

    if not level:
        return ""

    lookup_key = level.upper()

    if lookup_key in DEFAULT_LEVEL_ABBREVIATIONS:
        return DEFAULT_LEVEL_ABBREVIATIONS[lookup_key]

    first_word = re.split(r"\s+", level.strip())[0]

    if not first_word:
        return ""

    return first_word[0].upper()


def clean_match_name(raw_name: str) -> str:
    """
    Final match-name cleaner.

    Rules:
    - Remove R1, R2, R3, R4, R5, etc. anywhere in the match name.
    - Delete all lowercase letters.
    - Delete all symbols/spaces/punctuation.
    - Keep only numbers and capital letters.

    Examples:
    14CR1AM1 -> 14CAM1
    14Cr1aM1 -> 14CM1
    15PowerR2BM3 -> 15PBM3
    """
    if not raw_name:
        return ""

    text = normalize_power_text(str(raw_name))

    # Remove round labels like R1, R2, R10 anywhere in the name.
    text = re.sub(r"R\d+", "", text)

    # Keep only numeric and capital characters.
    text = re.sub(r"[^A-Z0-9]", "", text)

    return text


def match_age_key(match_code: str) -> str:
    """
    Age-only comparison key.

    For schedule-change checks, match names are intentionally compared only by
    the leading age code. This keeps the comparison focused on structure/age
    changes instead of pool, round, or individual match-number naming changes.
    """
    code = clean_match_name(match_code)
    return code[:2] if len(code) >= 2 else code


def match_age_label(match_code: str) -> str:
    age = match_age_key(match_code)

    if age and age.isdigit():
        return f"{age}U"

    return age or "Match"


def build_level_rows(divisions: list[dict]) -> list[dict]:
    level_map = {}

    for division in divisions:
        division_name = division.get("Name", "")
        level = division_level_name(division_name)

        if level not in level_map:
            level_map[level] = {
                "Division Level": level,
                "Abbreviation": default_level_abbreviation(level),
                "Divisions Using This Level": [],
            }

        level_map[level]["Divisions Using This Level"].append(division_name)

    rows = []

    for level in sorted(level_map.keys()):
        row = level_map[level]
        row["Divisions Using This Level"] = ", ".join(row["Divisions Using This Level"])
        rows.append(row)

    return rows


def resolve_level_abbreviations(edited_level_rows: list[dict]) -> dict:
    level_abbreviations = {}

    for row in edited_level_rows:
        level = normalize_level_name(str(row.get("Division Level", "")).strip())
        abbreviation = normalize_power_text(str(row.get("Abbreviation", "")).strip())

        if not abbreviation:
            abbreviation = default_level_abbreviation(level)

        # Abbreviations should only use capital letters/numbers.
        abbreviation = re.sub(r"[^A-Z0-9]", "", abbreviation)

        level_abbreviations[level] = abbreviation

    return level_abbreviations


def build_division_settings_from_levels(divisions: list[dict], level_abbreviations: dict) -> dict:
    division_settings = {}

    for division in divisions:
        division_id = division.get("DivisionId")
        division_name = division.get("Name", "")
        age = division_age(division_name)
        level = division_level_name(division_name)

        level_abbreviation = level_abbreviations.get(
            level,
            default_level_abbreviation(level),
        )

        final_abbreviation = f"{age}{level_abbreviation}" if age else level_abbreviation

        division_settings[division_id] = {
            "abbreviation": clean_match_name(final_abbreviation),
            "color": display_hex_color(division.get("ColorHex")),
        }

    return division_settings


def get_division_correction(match: dict, division_settings: dict) -> dict:
    division = match.get("Division", {}) or {}
    division_id = division.get("DivisionId")

    if division_id in division_settings:
        return division_settings[division_id]

    division_name = division.get("Name", "")
    age = division_age(division_name)
    level = division_level_name(division_name)
    level_abbreviation = default_level_abbreviation(level)
    final_abbreviation = f"{age}{level_abbreviation}" if age else level_abbreviation

    return {
        "abbreviation": clean_match_name(final_abbreviation),
        "color": display_hex_color(division.get("ColorHex")),
    }


def build_assignment_match_display_name(match: dict, division_settings: dict) -> str:
    correction = get_division_correction(match, division_settings)
    division_code = correction["abbreviation"]

    short_name = match.get("CompleteShortName") or ""

    return clean_match_name(f"{division_code}{short_name}")


def build_worksheet_match_display_name(match: dict, division_settings: dict) -> str:
    correction = get_division_correction(match, division_settings)
    division_code = correction["abbreviation"]

    short_name = match.get("CompleteShortName") or ""

    return clean_match_name(f"{division_code}{short_name}")


# =========================
# WORKBOOK STYLE HELPERS
# =========================

def compact_event_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", text or "AESEvent")


def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)

    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)

    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)


def bolded_font_from_cell(cell):
    existing = copy(cell.font)
    existing.bold = True
    return existing


def load_template_sheet_and_theme(template_path: str | None):
    if not template_path:
        return None, None

    path = Path(template_path)

    if not path.exists():
        return None, None

    template_wb = load_workbook(path)
    template_ws = template_wb[template_wb.sheetnames[0]]
    template_theme = template_wb.loaded_theme

    return template_ws, template_theme


def template_row_for_output_row(output_row: int) -> int:
    if output_row == 1:
        return 1

    return 2 + ((output_row - 2) % len(ROLE_ROWS))


def clear_unused_match_cell_fills(ws, start_row: int, max_row: int, start_col: int, max_col: int):
    if max_col < start_col:
        return

    for row in range(start_row, max_row + 1):
        for col in range(start_col, max_col + 1):
            cell = ws.cell(row=row, column=col)

            if cell.value is None or str(cell.value).strip() == "":
                cell.fill = PatternFill(fill_type=None)


def apply_template_columns_a_to_d(ws, template_ws, max_row: int):
    if template_ws is None:
        apply_builtin_columns_a_to_d_style(ws, max_row)
        return

    for col in range(1, 5):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = template_ws.column_dimensions[letter].width

    for row in range(1, max_row + 1):
        template_row = template_row_for_output_row(row)

        if template_ws.row_dimensions[template_row].height is not None:
            ws.row_dimensions[row].height = template_ws.row_dimensions[template_row].height

        for col in range(1, 5):
            copy_cell_style(
                template_ws.cell(row=template_row, column=col),
                ws.cell(row=row, column=col),
            )

    for row in range(1, max_row + 1):
        c_cell = ws.cell(row=row, column=3)
        c_cell.number_format = "h:mm AM/PM"

        if c_cell.value is not None:
            c_cell.font = bolded_font_from_cell(c_cell)

    d1 = ws["D1"]
    d1.value = None
    d1.fill = PatternFill(fill_type=None)


def apply_builtin_columns_a_to_d_style(ws, max_row: int):
    fallback_blue = "FFD9EAF7"

    widths = {
        "A": 10.15625,
        "B": 12.41796875,
        "C": 8.578125,
        "D": 7.0,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in range(1, max_row + 1):
        for col in range(1, 5):
            cell = ws.cell(row=row, column=col)
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        c_cell = ws.cell(row=row, column=3)
        c_cell.number_format = "h:mm AM/PM"

        if c_cell.value is not None:
            c_cell.font = Font(name="Calibri", size=11, bold=True)

        d_cell = ws.cell(row=row, column=4)
        d_cell.fill = PatternFill(fill_type="solid", fgColor=fallback_blue)
        d_cell.font = Font(name="Calibri", size=11, bold=True)
        d_cell.alignment = Alignment(horizontal="center", vertical="center")

    d1 = ws["D1"]
    d1.value = None
    d1.fill = PatternFill(fill_type=None)


def apply_template_court_header_style(ws, template_ws, start_col: int, end_col: int):
    if template_ws is None:
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(col)].width = 15.68359375
        return

    source_cell = template_ws["E1"]
    source_width = template_ws.column_dimensions["E"].width

    for col in range(start_col, end_col + 1):
        copy_cell_style(source_cell, ws.cell(row=1, column=col))
        ws.column_dimensions[get_column_letter(col)].width = source_width


def apply_template_match_area_style(ws, template_ws, start_col: int, end_col: int, max_row: int):
    if template_ws is None:
        for row in range(2, max_row + 1):
            for col in range(start_col, end_col + 1):
                ws.cell(row=row, column=col).font = Font(name="Calibri", size=11)
                ws.cell(row=row, column=col).alignment = Alignment(
                    horizontal="center", vertical="center"
                )
        return

    for row in range(2, max_row + 1):
        template_row = template_row_for_output_row(row)

        for col in range(start_col, end_col + 1):
            copy_cell_style(
                template_ws.cell(row=template_row, column=5),
                ws.cell(row=row, column=col),
            )


# =========================
# DATA SHAPING
# =========================

def flatten_matches(match_data: dict, tz_name: str):
    court_schedules = match_data.get("CourtSchedules", [])

    matches = []
    court_names = []

    for court in court_schedules:
        court_name = court.get("Name", "")
        court_names.append(court_name)

        for match in court.get("CourtMatches", []):
            matches.append(
                {
                    "court_name": court_name,
                    "match": match,
                    "start_dt": ms_to_local_datetime(
                        match["ScheduledStartDateTime"], tz_name
                    ),
                    "end_dt": ms_to_local_datetime(
                        match["ScheduledEndDateTime"], tz_name
                    ),
                }
            )

    return court_names, matches


def get_grid_start_end(sheet_date: date, matches: list, tz_name: str):
    if matches:
        earliest = min(item["start_dt"] for item in matches)
        latest = max(item["end_dt"] for item in matches)

        grid_start = earliest.replace(second=0, microsecond=0)
        grid_end = round_up_to_half_hour(latest)
    else:
        grid_start = datetime(
            sheet_date.year,
            sheet_date.month,
            sheet_date.day,
            8,
            0,
            tzinfo=ZoneInfo(tz_name),
        )
        grid_end = grid_start + timedelta(hours=12)

    return grid_start, grid_end


def build_grid_times(grid_start: datetime, grid_end: datetime) -> list[datetime]:
    grid_times = []
    current_time = grid_start

    while current_time <= grid_end:
        grid_times.append(current_time.replace(second=0, microsecond=0))
        current_time += timedelta(minutes=30)

    return grid_times


def find_time_row_for_match(match_start_dt: datetime, time_lookup: dict):
    clean_start = match_start_dt.replace(second=0, microsecond=0)

    if clean_start in time_lookup:
        return time_lookup[clean_start]

    grid_times = sorted(time_lookup.keys())
    target_time = round_match_start_to_next_grid_time(clean_start, grid_times)

    if target_time is None:
        return None

    return time_lookup[target_time]


# =========================
# ASSIGNMENT GRID WORKBOOK
# =========================

def build_assignment_time_grid(ws, sheet_date: date, grid_start: datetime, grid_end: datetime):
    ws["A1"] = "Notes"
    ws["B1"] = "Notes"
    ws["C1"] = c1_label_for_date(sheet_date)
    ws["D1"] = None

    time_to_base_row = {}
    grid_times = build_grid_times(grid_start, grid_end)

    row = 2
    for grid_time in grid_times:
        time_to_base_row[grid_time] = row

        time_cell = ws.cell(row=row, column=3)
        time_cell.value = grid_time.time()
        time_cell.number_format = "h:mm AM/PM"

        for offset, label in enumerate(ROLE_ROWS):
            ws.cell(row=row + offset, column=4).value = label

        row += len(ROLE_ROWS)

    return time_to_base_row, row - 1


def build_assignment_grid_sheet(
    ws,
    sheet_date: date,
    match_data: dict,
    tz_name: str,
    division_settings: dict,
    template_ws=None,
):
    court_names, matches = flatten_matches(match_data, tz_name)
    grid_start, grid_end = get_grid_start_end(sheet_date, matches, tz_name)

    time_to_base_row, max_row = build_assignment_time_grid(
        ws,
        sheet_date,
        grid_start,
        grid_end,
    )

    start_court_col = 5
    end_court_col = start_court_col + len(court_names) - 1

    apply_template_columns_a_to_d(ws, template_ws, max_row)

    if court_names:
        apply_template_court_header_style(ws, template_ws, start_court_col, end_court_col)
        apply_template_match_area_style(ws, template_ws, start_court_col, end_court_col, max_row)

    for idx, court_name in enumerate(court_names):
        col = start_court_col + idx
        ws.cell(row=1, column=col).value = court_name

    court_to_col = {
        court_name: start_court_col + idx for idx, court_name in enumerate(court_names)
    }

    for item in matches:
        court_name = item["court_name"]
        match = item["match"]

        base_row = find_time_row_for_match(item["start_dt"], time_to_base_row)

        if base_row is None:
            continue

        col = court_to_col[court_name]

        match_cell = ws.cell(row=base_row, column=col)
        format_cell = ws.cell(row=base_row + 1, column=col)

        match_name = build_assignment_match_display_name(match, division_settings)

        if match_name:
            match_cell.value = match_name
            match_cell.font = Font(name="Calibri", size=11, bold=True)
            match_cell.alignment = Alignment(horizontal="center", vertical="center")

            correction = get_division_correction(match, division_settings)
            fill_color = normalize_hex_color(correction["color"])
            match_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

            format_cell.value = determine_match_format(match)
            format_cell.font = Font(name="Calibri", size=11, bold=False)
            format_cell.alignment = Alignment(horizontal="center", vertical="center")

    if court_names:
        clear_unused_match_cell_fills(
            ws=ws,
            start_row=2,
            max_row=max_row,
            start_col=start_court_col,
            max_col=end_court_col,
        )

    ws["D1"].value = None
    ws["D1"].fill = PatternFill(fill_type=None)
    ws.freeze_panes = None


# =========================
# WORKSHEET GRID WORKBOOK
# =========================

def build_worksheet_grid_sheet(
    ws,
    sheet_date: date,
    match_data: dict,
    tz_name: str,
    division_settings: dict,
):
    court_names, matches = flatten_matches(match_data, tz_name)
    grid_start, grid_end = get_grid_start_end(sheet_date, matches, tz_name)

    ws["A1"] = c1_label_for_date(sheet_date)
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    start_court_col = 2

    for idx, court_name in enumerate(court_names):
        col = start_court_col + idx
        cell = ws.cell(row=1, column=col)
        cell.value = court_name
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    time_to_row = {}
    grid_times = build_grid_times(grid_start, grid_end)

    row = 2
    for grid_time in grid_times:
        time_to_row[grid_time] = row

        time_cell = ws.cell(row=row, column=1)
        time_cell.value = grid_time.time()
        time_cell.number_format = "h:mm AM/PM"
        time_cell.font = Font(name="Calibri", size=11, bold=True)
        time_cell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1

    max_row = row - 1
    max_col = start_court_col + len(court_names) - 1

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Calibri", size=11, bold=(r == 1 or c == 1))

    court_to_col = {
        court_name: start_court_col + idx for idx, court_name in enumerate(court_names)
    }

    for item in matches:
        court_name = item["court_name"]
        match = item["match"]

        row = find_time_row_for_match(item["start_dt"], time_to_row)

        if row is None:
            continue

        col = court_to_col[court_name]
        cell = ws.cell(row=row, column=col)

        match_name = build_worksheet_match_display_name(match, division_settings)

        if match_name:
            cell.value = match_name
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

            correction = get_division_correction(match, division_settings)
            fill_color = normalize_hex_color(correction["color"])
            cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    if court_names:
        clear_unused_match_cell_fills(
            ws=ws,
            start_row=2,
            max_row=max_row,
            start_col=start_court_col,
            max_col=max_col,
        )

    ws.column_dimensions["A"].width = 10

    for col in range(start_court_col, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15

    for r in range(1, max_row + 1):
        ws.row_dimensions[r].height = 18

    ws.freeze_panes = "B2"




# =========================
# GRID COMPARISON HELPERS
# =========================

def is_blank_cell_value(value) -> bool:
    return value is None or str(value).strip() == ""


def parse_grid_time_value(value) -> time | None:
    """Parse a worksheet time cell only when it is truly a time value."""
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)

    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)

    if isinstance(value, (int, float)):
        # Excel stores times as fractional days. Ignore larger values because
        # they are usually dates/counts in this type of workbook.
        if 0 <= value < 1:
            total_minutes = round(value * 24 * 60)
            hour = (total_minutes // 60) % 24
            minute = total_minutes % 60
            return time(hour=hour, minute=minute)
        return None

    if isinstance(value, str):
        text = value.strip()
        match = re.fullmatch(r"(\d{1,2}):(\d{2})(?:\s*([AP]M))?", text, flags=re.IGNORECASE)

        if not match:
            return None

        hour = int(match.group(1))
        minute = int(match.group(2))
        suffix = match.group(3).upper() if match.group(3) else None

        if minute > 59 or hour > 23:
            return None

        if suffix:
            if hour < 1 or hour > 12:
                return None
            if suffix == "PM" and hour != 12:
                hour += 12
            if suffix == "AM" and hour == 12:
                hour = 0

        return time(hour=hour, minute=minute)

    return None


def format_time_key(t: time) -> str:
    return f"{t.hour:02d}:{t.minute:02d}"


def format_time_label(t: time) -> str:
    hour = t.hour % 12 or 12
    suffix = "AM" if t.hour < 12 else "PM"
    return f"{hour}:{t.minute:02d} {suffix}"


def schedule_header_label(value) -> bool:
    if not isinstance(value, str):
        return False

    text = value.strip().upper()
    return bool(
        re.fullmatch(r"[A-Z]{2}\s+\d{1,2}/\d{1,2}", text)
        or re.fullmatch(r"[A-Z]{2}\s+\d{1,2}-\d{1,2}", text)
    )


def find_schedule_header_row(ws) -> int | None:
    """
    Detects the row containing the date/court headers.

    Compatible with:
    - Generated worksheet grids where the schedule starts on row 1.
    - Workbooks with assignment/name rows above the schedule grid.
    """
    max_scan_row = min(ws.max_row, 75)

    for row in range(1, max_scan_row + 1):
        header_values = [ws.cell(row=row, column=col).value for col in range(2, ws.max_column + 1)]
        non_empty_headers = [value for value in header_values if not is_blank_cell_value(value)]
        text_header_count = sum(1 for value in non_empty_headers if isinstance(value, str) and value.strip())

        if len(non_empty_headers) < 3:
            continue

        # Most compatible grids have court names as text. If the date label in
        # column A is clear, allow numeric court names too.
        if text_header_count < 3 and not schedule_header_label(ws.cell(row=row, column=1).value):
            continue

        time_offsets_below = []
        for check_row in range(row + 1, min(ws.max_row, row + 6) + 1):
            if parse_grid_time_value(ws.cell(row=check_row, column=1).value) is not None:
                time_offsets_below.append(check_row - row)

        if len(time_offsets_below) >= 2 and min(time_offsets_below) <= 2:
            return row

    return None


def normalize_uploaded_match_value(value) -> str:
    if value is None or isinstance(value, (datetime, date, time, int, float)):
        return ""

    text = str(value).strip()

    if not text or parse_grid_time_value(text) is not None:
        return ""

    match_code = clean_match_name(text)

    if len(match_code) < 3:
        return ""

    if not re.search(r"\d", match_code):
        return ""

    return match_code


def normalize_court_exact_key(court_name: str) -> str:
    clean = re.sub(r"[^A-Z0-9]", "", str(court_name or "").upper())
    return clean or "COURT"


def court_number(court_name: str) -> int | None:
    match = re.search(r"(\d+)\s*$", str(court_name or "").strip())
    return int(match.group(1)) if match else None


def make_schedule_entry(
    court_name: str,
    match_time: time,
    match_code: str,
    source: str,
    division_label: str = "",
):
    return {
        "court_name": str(court_name or "").strip(),
        "court_key": normalize_court_exact_key(court_name),
        "court_number": court_number(court_name),
        "time_key": format_time_key(match_time),
        "time_label": format_time_label(match_time),
        "match_code": clean_match_name(match_code),
        "match_age": match_age_key(match_code),
        "match_age_label": match_age_label(match_code),
        "source": source,
        "division_label": division_label or describe_match_code(match_code),
    }


def get_uploaded_workbook_sheet_names(uploaded_file) -> list[str]:
    uploaded_file.seek(0)
    wb = load_workbook(uploaded_file, read_only=True, data_only=True)
    return wb.sheetnames


def default_uploaded_sheet_index(sheet_names: list[str], selected_date: date) -> int:
    if not sheet_names:
        return 0

    expected_names = {
        sheet_name_for_date(selected_date).upper(),
        c1_label_for_date(selected_date).upper(),
    }

    for idx, sheet_name in enumerate(sheet_names):
        if sheet_name.upper() in expected_names:
            return idx

    return 0


def parse_uploaded_workbook_schedule(uploaded_file, sheet_name: str | None = None) -> dict:
    uploaded_file.seek(0)
    wb = load_workbook(uploaded_file, data_only=True)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb[wb.sheetnames[0]]

    header_row = find_schedule_header_row(ws)

    if header_row is None:
        raise ValueError(
            "Could not find a schedule grid. Make sure the uploaded workbook has court names across the top and times down column A."
        )

    court_columns = []

    for col in range(2, ws.max_column + 1):
        value = ws.cell(row=header_row, column=col).value

        if is_blank_cell_value(value):
            continue

        court_columns.append((col, str(value).strip()))

    entries = []

    for row in range(header_row + 1, ws.max_row + 1):
        parsed_time = parse_grid_time_value(ws.cell(row=row, column=1).value)

        if parsed_time is None:
            continue

        for col, court_name in court_columns:
            match_code = normalize_uploaded_match_value(ws.cell(row=row, column=col).value)

            if match_code:
                entries.append(
                    make_schedule_entry(
                        court_name=court_name,
                        match_time=parsed_time,
                        match_code=match_code,
                        source="uploaded",
                    )
                )

    return {
        "sheet_name": ws.title,
        "header_row": header_row,
        "court_count": len(court_columns),
        "match_count": len(entries),
        "entries": entries,
    }


def grid_time_for_online_match(match_start_dt: datetime, grid_times: list[datetime]) -> time:
    rounded_dt = round_match_start_to_next_grid_time(match_start_dt, grid_times)

    if rounded_dt is None:
        rounded_dt = round_up_to_half_hour(match_start_dt)

    return rounded_dt.time().replace(second=0, microsecond=0)


def build_online_schedule_entries(match_data: dict, tz_name: str, division_settings: dict) -> list[dict]:
    court_names, matches = flatten_matches(match_data, tz_name)

    if matches:
        sheet_date = matches[0]["start_dt"].date()
    else:
        sheet_date = datetime.now(ZoneInfo(tz_name)).date()

    grid_start, grid_end = get_grid_start_end(sheet_date, matches, tz_name)
    grid_times = build_grid_times(grid_start, grid_end)

    entries = []

    for item in matches:
        match = item["match"]
        match_code = build_worksheet_match_display_name(match, division_settings)

        if not match_code:
            continue

        division = match.get("Division", {}) or {}
        division_label = division.get("Name", "") or describe_match_code(match_code)
        match_time = grid_time_for_online_match(item["start_dt"], grid_times)

        entries.append(
            make_schedule_entry(
                court_name=item["court_name"],
                match_time=match_time,
                match_code=match_code,
                source="online",
                division_label=division_label,
            )
        )

    return entries


def describe_match_code(match_code: str) -> str:
    """Best-effort human label for workbook-only match codes."""
    code = clean_match_name(match_code)
    age_match = re.match(r"^(\d{1,2})", code)

    if not age_match:
        return code

    age = age_match.group(1)
    remainder = code[len(age):]

    abbreviation_names = {
        "AS": "Aspire",
        "SP": "Spirit",
        "CL": "Classic",
        "O": "Open",
        "P": "Premier",
        "E": "Elite",
        "S": "Select",
        "A": "Ascend",
        "C": "Club",
        "G": "Gold",
        "B": "Bronze",
        "D": "Diamond",
        "R": "Regional",
        "N": "National",
        "U": "USA",
    }

    for abbreviation in sorted(abbreviation_names.keys(), key=len, reverse=True):
        if remainder.startswith(abbreviation):
            return f"{age} {abbreviation_names[abbreviation]}"

    return code


def build_court_aliases(uploaded_entries: list[dict], online_entries: list[dict]) -> tuple[dict, dict, list[str]]:
    """
    Exact court names are preferred. If exact names differ, unique trailing court
    numbers are used as a fallback so OCCC 1 can compare to North 1 when needed.
    """
    uploaded_courts = {}
    online_courts = {}

    for entry in uploaded_entries:
        uploaded_courts.setdefault(entry["court_key"], entry["court_name"])

    for entry in online_entries:
        online_courts.setdefault(entry["court_key"], entry["court_name"])

    aliases = {key: key for key in online_courts.keys()}
    canonical_names = {key: name for key, name in online_courts.items()}
    notes = []

    def unique_number_map(court_lookup: dict) -> dict:
        grouped = defaultdict(list)

        for key, name in court_lookup.items():
            number = court_number(name)
            if number is not None:
                grouped[number].append((key, name))

        return {
            number: values[0]
            for number, values in grouped.items()
            if len(values) == 1
        }

    uploaded_by_number = unique_number_map(uploaded_courts)
    online_by_number = unique_number_map(online_courts)

    for uploaded_key, uploaded_name in uploaded_courts.items():
        if uploaded_key in online_courts:
            aliases[uploaded_key] = uploaded_key
            canonical_names.setdefault(uploaded_key, online_courts[uploaded_key])
            continue

        number = court_number(uploaded_name)

        if number is not None and number in uploaded_by_number and number in online_by_number:
            online_key, online_name = online_by_number[number]
            aliases[uploaded_key] = online_key
            canonical_names[online_key] = online_name
            notes.append(f"Matched uploaded court '{uploaded_name}' to online court '{online_name}' by court number.")
        else:
            aliases[uploaded_key] = uploaded_key
            canonical_names.setdefault(uploaded_key, uploaded_name)

    return aliases, canonical_names, notes


def aligned_position(entry: dict, court_aliases: dict) -> tuple[str, str]:
    return (court_aliases.get(entry["court_key"], entry["court_key"]), entry["time_key"])


def entry_location_text(entry: dict, canonical_court_names: dict | None = None) -> str:
    if canonical_court_names and "aligned_court_key" in entry:
        court_name = canonical_court_names.get(entry["aligned_court_key"], entry["court_name"])
    else:
        court_name = entry["court_name"]

    return f"{court_name} at {entry['time_label']}"


def entries_by_position(entries: list[dict], court_aliases: dict) -> dict:
    lookup = {}

    for entry in entries:
        aligned_key, time_key = aligned_position(entry, court_aliases)
        entry["aligned_court_key"] = aligned_key
        lookup[(aligned_key, time_key)] = entry

    return lookup


def entries_by_match_code(entries: list[dict]) -> dict:
    lookup = defaultdict(list)

    for entry in entries:
        lookup[entry["match_code"]].append(entry)

    return lookup


def entries_by_match_age(entries: list[dict]) -> dict:
    lookup = defaultdict(list)

    for entry in entries:
        lookup[entry["match_age"]].append(entry)

    return lookup


def entry_aligned_court_key(entry: dict, court_aliases: dict | None = None) -> str:
    if "aligned_court_key" in entry:
        return entry["aligned_court_key"]

    if court_aliases:
        return court_aliases.get(entry["court_key"], entry["court_key"])

    return entry["court_key"]


def collect_changed_court_keys(*entry_groups) -> set[str]:
    changed_keys = set()

    for group in entry_groups:
        if not group:
            continue

        if isinstance(group, dict):
            group = group.values()

        for item in group:
            if isinstance(item, dict) and "old" in item and "new" in item:
                changed_keys.add(entry_aligned_court_key(item["old"]))
                changed_keys.add(entry_aligned_court_key(item["new"]))
            elif isinstance(item, dict) and "old_entries" in item and "new_entries" in item:
                for entry in item["old_entries"] + item["new_entries"]:
                    changed_keys.add(entry_aligned_court_key(entry))
            elif isinstance(item, dict) and "court_key" in item:
                changed_keys.add(entry_aligned_court_key(item))
            elif isinstance(item, list):
                for entry in item:
                    if isinstance(entry, dict) and "court_key" in entry:
                        changed_keys.add(entry_aligned_court_key(entry))

    return changed_keys


def pair_age_moves(removed_candidates: list[dict], added_candidates: list[dict]) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Pair possible moves using age only.

    Because the comparison intentionally ignores full match names, this treats a
    removed 18U position and an added 18U position as a likely move. When there
    are multiple of the same age, entries are paired by time/court sort order.
    """
    removed_by_age = defaultdict(list)
    added_by_age = defaultdict(list)

    for entry in removed_candidates:
        removed_by_age[entry["match_age"]].append(entry)

    for entry in added_candidates:
        added_by_age[entry["match_age"]].append(entry)

    moved_matches = []
    remaining_removed = []
    remaining_added = []

    all_ages = sorted(set(removed_by_age) | set(added_by_age), key=natural_sort_key)

    for age in all_ages:
        old_entries = sorted_entries_by_time(removed_by_age.get(age, []))
        new_entries = sorted_entries_by_time(added_by_age.get(age, []))
        pair_count = min(len(old_entries), len(new_entries))

        for idx in range(pair_count):
            old_entry = old_entries[idx]
            new_entry = new_entries[idx]
            moved_matches.append({
                "match_age": age,
                "match_age_label": new_entry.get("match_age_label") or old_entry.get("match_age_label") or match_age_label(age),
                "old_entries": [old_entry],
                "new_entries": [new_entry],
                "division_label": new_entry.get("match_age_label") or old_entry.get("match_age_label") or match_age_label(age),
            })

        remaining_removed.extend(old_entries[pair_count:])
        remaining_added.extend(new_entries[pair_count:])

    return moved_matches, remaining_removed, remaining_added


def compare_schedule_structures(uploaded_entries: list[dict], online_entries: list[dict]) -> dict:
    court_aliases, canonical_court_names, alias_notes = build_court_aliases(uploaded_entries, online_entries)

    uploaded_by_position = entries_by_position(uploaded_entries, court_aliases)
    online_by_position = entries_by_position(online_entries, court_aliases)

    uploaded_positions = set(uploaded_by_position.keys())
    online_positions = set(online_by_position.keys())

    changed_positions = []

    for position in sorted(uploaded_positions & online_positions):
        uploaded_entry = uploaded_by_position[position]
        online_entry = online_by_position[position]

        # Match-name comparisons are age-only by request.
        if uploaded_entry["match_age"] != online_entry["match_age"]:
            changed_positions.append({
                "position": position,
                "old": uploaded_entry,
                "new": online_entry,
            })

    removed_candidates = [
        uploaded_by_position[position]
        for position in sorted(uploaded_positions - online_positions)
    ]

    added_candidates = [
        online_by_position[position]
        for position in sorted(online_positions - uploaded_positions)
    ]

    moved_matches, removed_entries, added_entries = pair_age_moves(
        removed_candidates,
        added_candidates,
    )

    added_by_court = group_change_entries_by_court(added_entries, canonical_court_names)
    removed_by_court = group_change_entries_by_court(removed_entries, canonical_court_names)

    changed_court_keys = collect_changed_court_keys(
        added_entries,
        removed_entries,
        moved_matches,
        changed_positions,
    )

    is_unchanged = not any([
        added_entries,
        removed_entries,
        moved_matches,
        changed_positions,
    ])

    changed_court_names = [
        canonical_court_names.get(key, key)
        for key in sorted(changed_court_keys, key=lambda value: natural_sort_key(canonical_court_names.get(value, value)))
    ]

    return {
        "is_unchanged": is_unchanged,
        "added_entries": added_entries,
        "removed_entries": removed_entries,
        "added_by_court": added_by_court,
        "removed_by_court": removed_by_court,
        "moved_matches": moved_matches,
        "changed_positions": changed_positions,
        "changed_court_keys": changed_court_keys,
        "changed_court_names": changed_court_names,
        "canonical_court_names": canonical_court_names,
        "alias_notes": alias_notes,
        "uploaded_match_count": len(uploaded_entries),
        "online_match_count": len(online_entries),
    }


def format_time_list(entries: list[dict]) -> str:
    times = []
    seen = set()

    for entry in sorted_entries_by_time(entries):
        if entry["time_key"] not in seen:
            seen.add(entry["time_key"])
            times.append(entry["time_label"])

    return ", ".join(times)


def render_grouped_change_lines(grouped_entries: dict, verb: str):
    for court_name, entries in grouped_entries.items():
        count = len(entries)
        plural = "match" if count == 1 else "matches"
        st.write(f"- **{court_name}** {verb} {count} {plural} ({format_time_list(entries)}).")


def render_comparison_results(result: dict, uploaded_summary: dict, selected_date: date):
    st.caption(
        f"Compared {uploaded_summary['match_count']} uploaded matches from '{uploaded_summary['sheet_name']}' "
        f"to {result['online_match_count']} online matches for {format_long_date(selected_date)}. "
        "Match-name differences are compared by age only using the first two characters of the cleaned match name."
    )

    if result["alias_notes"]:
        with st.expander("Court-name matching notes"):
            st.write(f"Matched {len(result['alias_notes'])} uploaded courts to online courts by court number.")
            for note in result["alias_notes"][:20]:
                st.write(f"- {note}")
            if len(result["alias_notes"]) > 20:
                st.write(f"- ...and {len(result['alias_notes']) - 20} more.")

    if result["is_unchanged"]:
        st.success("✅ Grid Structure Verified & Unchanged")
        return

    st.warning("Grid structure changes found.")

    if result["removed_by_court"]:
        st.markdown("**Removed matches**")
        render_grouped_change_lines(result["removed_by_court"], "removed")

    if result["added_by_court"]:
        st.markdown("**Added matches**")
        render_grouped_change_lines(result["added_by_court"], "added")

    if result["moved_matches"]:
        st.markdown("**Moved matches**")
        for move in result["moved_matches"]:
            old_locations = ", ".join(
                entry_location_text(entry, result["canonical_court_names"])
                for entry in sorted_entries_by_time(move["old_entries"])
            )
            new_locations = ", ".join(
                entry_location_text(entry, result["canonical_court_names"])
                for entry in sorted_entries_by_time(move["new_entries"])
            )
            st.write(
                f"- **{move['match_age_label']}** moved from {old_locations} to {new_locations}."
            )

    if result["changed_positions"]:
        st.markdown("**Changed age at same court and time**")
        for change in result["changed_positions"]:
            old_entry = change["old"]
            new_entry = change["new"]
            court_name = result["canonical_court_names"].get(
                new_entry.get("aligned_court_key"),
                new_entry["court_name"],
            )
            st.write(
                f"- **{court_name} at {new_entry['time_label']}** is now "
                f"**{new_entry['match_age_label']}**; was **{old_entry['match_age_label']}**."
            )


# =========================
# CREATE WORKBOOKS IN MEMORY
# =========================

def save_workbook_to_bytes(wb: Workbook) -> bytes:
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def create_workbooks(event_key: str, event_info: dict, division_settings: dict):
    event_name = event_info.get("Name", "AESEvent")

    start_date = parse_event_date(event_info["StartDate"])
    end_date = parse_event_date(event_info["EndDate"])

    template_ws, template_theme = load_template_sheet_and_theme(TEMPLATE_PATH)

    assignment_wb = Workbook()
    assignment_wb.remove(assignment_wb.active)

    worksheet_wb = Workbook()
    worksheet_wb.remove(worksheet_wb.active)

    if template_theme:
        assignment_wb.loaded_theme = template_theme
        worksheet_wb.loaded_theme = template_theme

    for event_date in date_range(start_date, end_date):
        sheet_name = sheet_name_for_date(event_date)
        match_data = get_match_info(event_key, event_date)

        assignment_ws = assignment_wb.create_sheet(title=sheet_name)
        build_assignment_grid_sheet(
            ws=assignment_ws,
            sheet_date=event_date,
            match_data=match_data,
            tz_name=TIMEZONE,
            division_settings=division_settings,
            template_ws=template_ws,
        )

        worksheet_ws = worksheet_wb.create_sheet(title=sheet_name)
        build_worksheet_grid_sheet(
            ws=worksheet_ws,
            sheet_date=event_date,
            match_data=match_data,
            tz_name=TIMEZONE,
            division_settings=division_settings,
        )

    event_name_no_spaces = compact_event_name(event_name)
    suffix = timestamp_suffix(TIMEZONE)

    assignment_filename = f"AssignmentGrid_{event_name_no_spaces}{suffix}.xlsx"
    worksheet_filename = f"WorksheetGrid_{event_name_no_spaces}{suffix}.xlsx"

    return {
        "assignment_filename": assignment_filename,
        "assignment_bytes": save_workbook_to_bytes(assignment_wb),
        "worksheet_filename": worksheet_filename,
        "worksheet_bytes": save_workbook_to_bytes(worksheet_wb),
    }


# =========================
# MATCH COUNTS / CHANGED-COURT WORKBOOKS
# =========================

def count_matches_in_match_data(match_data: dict) -> int:
    return sum(
        len(court.get("CourtMatches", []) or [])
        for court in match_data.get("CourtSchedules", []) or []
    )


def get_event_match_counts_by_day(event_key: str, event_info: dict) -> dict:
    start_date = parse_event_date(event_info["StartDate"])
    end_date = parse_event_date(event_info["EndDate"])

    rows = []
    total = 0

    for event_date in date_range(start_date, end_date):
        match_data = get_match_info(event_key, event_date)
        count = count_matches_in_match_data(match_data)
        total += count
        rows.append({
            "Date": format_long_date(event_date),
            "Sheet": sheet_name_for_date(event_date),
            "Match Count": count,
        })

    return {
        "rows": rows,
        "total": total,
    }


def filter_match_data_to_court_keys(match_data: dict, included_court_keys: set[str]) -> dict:
    included_court_keys = set(included_court_keys or [])

    filtered = copy(match_data)
    filtered_court_schedules = []

    for court in match_data.get("CourtSchedules", []) or []:
        court_key = normalize_court_exact_key(court.get("Name", ""))

        if court_key in included_court_keys:
            filtered_court_schedules.append(court)

    filtered["CourtSchedules"] = filtered_court_schedules
    return filtered


def create_changed_court_workbooks(
    event_info: dict,
    event_date: date,
    match_data: dict,
    division_settings: dict,
    changed_court_keys: set[str],
):
    if not changed_court_keys:
        return None

    event_name = event_info.get("Name", "AESEvent")
    event_name_no_spaces = compact_event_name(event_name)
    date_suffix = event_date.strftime("%m%d")
    suffix = timestamp_suffix(TIMEZONE)

    filtered_match_data = filter_match_data_to_court_keys(match_data, changed_court_keys)
    template_ws, template_theme = load_template_sheet_and_theme(TEMPLATE_PATH)

    assignment_wb = Workbook()
    assignment_wb.remove(assignment_wb.active)

    worksheet_wb = Workbook()
    worksheet_wb.remove(worksheet_wb.active)

    if template_theme:
        assignment_wb.loaded_theme = template_theme
        worksheet_wb.loaded_theme = template_theme

    sheet_name = sheet_name_for_date(event_date)

    assignment_ws = assignment_wb.create_sheet(title=sheet_name)
    build_assignment_grid_sheet(
        ws=assignment_ws,
        sheet_date=event_date,
        match_data=filtered_match_data,
        tz_name=TIMEZONE,
        division_settings=division_settings,
        template_ws=template_ws,
    )

    worksheet_ws = worksheet_wb.create_sheet(title=sheet_name)
    build_worksheet_grid_sheet(
        ws=worksheet_ws,
        sheet_date=event_date,
        match_data=filtered_match_data,
        tz_name=TIMEZONE,
        division_settings=division_settings,
    )

    return {
        "assignment_filename": f"ChangedCourts_AssignmentGrid_{event_name_no_spaces}_{date_suffix}{suffix}.xlsx",
        "assignment_bytes": save_workbook_to_bytes(assignment_wb),
        "worksheet_filename": f"ChangedCourts_WorksheetGrid_{event_name_no_spaces}_{date_suffix}{suffix}.xlsx",
        "worksheet_bytes": save_workbook_to_bytes(worksheet_wb),
        "court_count": len(filtered_match_data.get("CourtSchedules", []) or []),
    }


# =========================
# STREAMLIT APP
# =========================

st.set_page_config(
    page_title="AES Grid Workbook Generator",
    page_icon="🏐",
    layout="wide",
)

top_left, top_right = st.columns([5, 1])

with top_left:
    st.title("AES Grid Workbook Generator")

with top_right:
    st.button("Start Over", on_click=reset_app, use_container_width=True)

st.write(
    "Paste an AES event key or full AES schedule URL, review the event information, "
    "set division level abbreviations, then generate both Excel workbooks."
)

if "event_info" not in st.session_state:
    st.session_state.event_info = None

if "event_key" not in st.session_state:
    st.session_state.event_key = ""

if "generated_workbooks" not in st.session_state:
    st.session_state.generated_workbooks = None

if "level_rows" not in st.session_state:
    st.session_state.level_rows = None

if "event_match_counts" not in st.session_state:
    st.session_state.event_match_counts = None

if "comparison_result" not in st.session_state:
    st.session_state.comparison_result = None

if "comparison_uploaded_summary" not in st.session_state:
    st.session_state.comparison_uploaded_summary = None

if "comparison_selected_date" not in st.session_state:
    st.session_state.comparison_selected_date = None

if "changed_court_workbooks" not in st.session_state:
    st.session_state.changed_court_workbooks = None


event_input = st.text_input(
    "AES Event Key or Schedule URL",
    placeholder="Example: PTAwMDAwNDEzMjQ90 or https://results.advancedeventsystems.com/event/PTAwMDAwNDEzMjQ90/home",
)

fetch_clicked = st.button("Fetch Event", type="primary")

if fetch_clicked:
    clean_key = extract_event_key(event_input)

    if not clean_key:
        st.error("Enter an event key or URL.")
    else:
        try:
            with st.spinner("Loading..."):
                st.session_state.event_info = get_event_info(clean_key)
                st.session_state.event_key = clean_key
                st.session_state.generated_workbooks = None
                st.session_state.comparison_result = None
                st.session_state.comparison_uploaded_summary = None
                st.session_state.comparison_selected_date = None
                st.session_state.changed_court_workbooks = None
                st.session_state.level_rows = build_level_rows(
                    st.session_state.event_info.get("Divisions", [])
                )
                st.session_state.event_match_counts = get_event_match_counts_by_day(
                    clean_key,
                    st.session_state.event_info,
                )

            st.success("Loaded.")
        except Exception:
            st.session_state.event_info = None
            st.session_state.generated_workbooks = None
            st.session_state.level_rows = None
            st.session_state.event_match_counts = None
            st.session_state.comparison_result = None
            st.session_state.comparison_uploaded_summary = None
            st.session_state.comparison_selected_date = None
            st.session_state.changed_court_workbooks = None
            st.error("Could not load event.")


event_info = st.session_state.event_info

if event_info:
    event_name = event_info.get("Name", "")
    location = event_info.get("Location", "")
    start_date = parse_event_date(event_info["StartDate"])
    end_date = parse_event_date(event_info["EndDate"])

    st.subheader("Event Information")
    st.markdown(f"### {event_name}")

    c1, c2, c3, c4 = st.columns([1.3, 1.3, 2.8, 1.2])

    c1.metric("Start Date", format_long_date(start_date))
    c2.metric("End Date", format_long_date(end_date))
    c3.metric("Location", location or "Not listed")

    match_counts = st.session_state.event_match_counts
    if match_counts:
        c4.metric("Total Matches", match_counts["total"])
    else:
        c4.metric("Total Matches", "—")

    st.caption(f"AES event key: `{st.session_state.event_key}`")

    if match_counts and match_counts.get("rows"):
        with st.expander("Match count by day", expanded=True):
            st.table(match_counts["rows"])

    st.subheader("Division Level Abbreviations")

    st.write(
        "Set one abbreviation for each unique division level. "
        "The app will add the age automatically. For example, Classic → CL gives 11CL, 12CL, 13CL."
    )

    if st.session_state.level_rows is None:
        st.session_state.level_rows = build_level_rows(event_info.get("Divisions", []))

    edited_level_rows = st.data_editor(
        st.session_state.level_rows,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        column_order=[
            "Division Level",
            "Abbreviation",
            "Divisions Using This Level",
        ],
        column_config={
            "Division Level": st.column_config.TextColumn(
                "Division Level",
                disabled=True,
                width="medium",
            ),
            "Abbreviation": st.column_config.TextColumn(
                "Abbreviation",
                help="Example: Open → O, Premier → P, Classic → CL, Aspire → AS",
                width="small",
            ),
            "Divisions Using This Level": st.column_config.TextColumn(
                "Divisions Using This Level",
                disabled=True,
                width="large",
            ),
        },
        key="level_editor",
    )

    level_abbreviations = resolve_level_abbreviations(edited_level_rows)

    division_settings = build_division_settings_from_levels(
        divisions=event_info.get("Divisions", []),
        level_abbreviations=level_abbreviations,
    )

    st.subheader("Generate Workbooks")

    generate_clicked = st.button("Generate", type="primary")

    if generate_clicked:
        try:
            with st.spinner("Generating..."):
                st.session_state.generated_workbooks = create_workbooks(
                    event_key=st.session_state.event_key,
                    event_info=event_info,
                    division_settings=division_settings,
                )

            st.success("Done.")
        except Exception:
            st.session_state.generated_workbooks = None
            st.error("Could not generate workbooks.")

    generated = st.session_state.generated_workbooks

    if generated:
        st.markdown("#### Downloads")

        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                label="Download Assignment Grid",
                data=generated["assignment_bytes"],
                file_name=generated["assignment_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_assignment_grid",
            )

        with d2:
            st.download_button(
                label="Download Worksheet Grid",
                data=generated["worksheet_bytes"],
                file_name=generated["worksheet_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_worksheet_grid",
            )


    st.divider()
    st.subheader("Compare Uploaded Workbook to Online Schedule")
    st.write(
        "Upload a previously generated or manually edited grid workbook, choose the day/sheet, "
        "and compare its court/time match structure to the current AES online schedule."
    )

    available_compare_dates = list(date_range(start_date, end_date))

    selected_compare_date = st.selectbox(
        "Day to compare",
        options=available_compare_dates,
        format_func=format_long_date,
        key="compare_date",
    )

    uploaded_compare_workbook = st.file_uploader(
        "Upload workbook to compare",
        type=["xlsx", "xlsm"],
        key="compare_workbook",
    )

    selected_compare_sheet = None

    if uploaded_compare_workbook is not None:
        try:
            compare_sheet_names = get_uploaded_workbook_sheet_names(uploaded_compare_workbook)
            default_sheet_index = default_uploaded_sheet_index(
                compare_sheet_names,
                selected_compare_date,
            )

            selected_compare_sheet = st.selectbox(
                "Workbook sheet to read",
                options=compare_sheet_names,
                index=default_sheet_index,
                key="compare_sheet",
            )
        except Exception as exc:
            st.error(f"Could not read workbook sheets: {exc}")

    compare_clicked = st.button(
        "Compare Uploaded Workbook to Current Online Schedule",
        type="secondary",
        key="compare_schedule_button",
    )

    if compare_clicked:
        if uploaded_compare_workbook is None:
            st.error("Upload a workbook first.")
        elif not selected_compare_sheet:
            st.error("Select a worksheet from the uploaded workbook.")
        else:
            try:
                with st.spinner("Comparing workbook to current AES schedule..."):
                    uploaded_summary = parse_uploaded_workbook_schedule(
                        uploaded_compare_workbook,
                        selected_compare_sheet,
                    )
                    current_match_data = get_match_info(
                        st.session_state.event_key,
                        selected_compare_date,
                    )
                    online_entries = build_online_schedule_entries(
                        current_match_data,
                        TIMEZONE,
                        division_settings,
                    )
                    comparison_result = compare_schedule_structures(
                        uploaded_summary["entries"],
                        online_entries,
                    )

                    changed_workbooks = None
                    if comparison_result["changed_court_keys"]:
                        changed_workbooks = create_changed_court_workbooks(
                            event_info=event_info,
                            event_date=selected_compare_date,
                            match_data=current_match_data,
                            division_settings=division_settings,
                            changed_court_keys=comparison_result["changed_court_keys"],
                        )

                st.session_state.comparison_result = comparison_result
                st.session_state.comparison_uploaded_summary = uploaded_summary
                st.session_state.comparison_selected_date = selected_compare_date
                st.session_state.changed_court_workbooks = changed_workbooks
            except Exception as exc:
                st.session_state.comparison_result = None
                st.session_state.comparison_uploaded_summary = None
                st.session_state.comparison_selected_date = None
                st.session_state.changed_court_workbooks = None
                st.error(f"Could not compare schedules: {exc}")

    if (
        st.session_state.comparison_result
        and st.session_state.comparison_uploaded_summary
        and st.session_state.comparison_selected_date
    ):
        render_comparison_results(
            st.session_state.comparison_result,
            st.session_state.comparison_uploaded_summary,
            st.session_state.comparison_selected_date,
        )

        changed_workbooks = st.session_state.changed_court_workbooks
        if changed_workbooks:
            changed_names = st.session_state.comparison_result.get("changed_court_names", [])
            st.markdown("#### Changed-Court Downloads")
            st.caption(
                f"These files include only the {changed_workbooks['court_count']} court(s) with detected changes: "
                + ", ".join(changed_names)
            )

            cd1, cd2 = st.columns(2)

            with cd1:
                st.download_button(
                    label="Download Changed Courts Assignment Grid",
                    data=changed_workbooks["assignment_bytes"],
                    file_name=changed_workbooks["assignment_filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="download_changed_assignment_grid",
                )

            with cd2:
                st.download_button(
                    label="Download Changed Courts Worksheet Grid",
                    data=changed_workbooks["worksheet_bytes"],
                    file_name=changed_workbooks["worksheet_filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="download_changed_worksheet_grid",
                )

else:
    st.info("Enter an AES event key or full AES schedule URL to begin.")