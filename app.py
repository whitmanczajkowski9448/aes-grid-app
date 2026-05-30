import re
from copy import copy
from datetime import datetime, date, timedelta, timezone
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


# =========================
# INPUT HELPERS
# =========================

def extract_event_key(user_input: str) -> str:
    """
    Accepts either:
    - PTAwMDAwNDEzMjQ90
    - https://results.advancedeventsystems.com/event/PTAwMDAwNDEzMjQ90/home
    - https://results.advancedeventsystems.com/event/PTAwMDAwNDEzMjQ90/court-schedule
    - https://results.advancedeventsystems.com/api/event/PTAwMDAwNDEzMjQ90
    """
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
        "division_rows",
        "division_editor",
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


def division_level_name(division_name: str) -> str:
    """
    Groups divisions by name regardless of age.

    Examples:
    11 Classic -> Classic
    12 Classic -> Classic
    13 Open -> Open
    18 Power -> P
    """
    name = normalize_power_text((division_name or "").strip())
    name = re.sub(r"^\d{1,2}\s*", "", name).strip()

    return name or "Division"


def division_age(division_name: str) -> str:
    match = re.search(r"^\s*(\d{1,2})", division_name or "")
    return match.group(1) if match else ""


def abbreviation_suffix(abbreviation: str, age: str) -> str:
    abbreviation = normalize_power_text((abbreviation or "").strip())

    if age and abbreviation.startswith(age):
        return abbreviation[len(age):]

    return abbreviation


def auto_division_abbreviation(division_name: str, code_alias: str = "") -> str:
    division_name = normalize_power_text((division_name or "").strip())
    code_alias = normalize_power_text((code_alias or "").strip())

    if re.fullmatch(r"\d{1,2}[A-Z]+", code_alias):
        return code_alias

    match = re.search(r"(\d{1,2})\s*([A-Za-z]+)", division_name)

    if not match:
        return code_alias or ""

    age = match.group(1)
    level_word = match.group(2).strip()

    if level_word.lower() == "power":
        level_letter = "P"
    else:
        level_letter = level_word[0].upper()

    return f"{age}{level_letter}"


def remove_round_from_match_short_name(short_name: str) -> str:
    short_name = normalize_power_text((short_name or "").strip())
    return re.sub(r"^R\d+", "", short_name)


def get_division_correction(match: dict, division_settings: dict) -> dict:
    division = match.get("Division", {}) or {}
    division_id = division.get("DivisionId")

    if division_id in division_settings:
        return division_settings[division_id]

    return {
        "abbreviation": auto_division_abbreviation(
            division.get("Name", ""),
            division.get("CodeAlias", ""),
        ),
        "color": display_hex_color(division.get("ColorHex")),
    }


def build_assignment_match_display_name(match: dict, division_settings: dict) -> str:
    correction = get_division_correction(match, division_settings)
    division_code = normalize_power_text(correction["abbreviation"])

    short_name = match.get("CompleteShortName") or ""
    short_name = remove_round_from_match_short_name(short_name)

    return normalize_power_text(f"{division_code}{short_name}")


def build_worksheet_match_display_name(match: dict, division_settings: dict) -> str:
    correction = get_division_correction(match, division_settings)
    division_code = normalize_power_text(correction["abbreviation"])

    short_name = normalize_power_text((match.get("CompleteShortName") or "").strip())

    return normalize_power_text(f"{division_code}{short_name}")


def build_division_rows(divisions: list[dict]) -> list[dict]:
    rows = []

    for division in divisions:
        name = division.get("Name", "")
        age = division_age(name)
        level = division_level_name(name)
        auto_abbrev = auto_division_abbreviation(name, division.get("CodeAlias", ""))

        rows.append(
            {
                "DivisionId": division.get("DivisionId"),
                "Division Name": name,
                "Level Group": level,
                "Team Count": division.get("TeamCount", ""),
                "AES Code": division.get("CodeAlias", ""),
                "Final Abbreviation": auto_abbrev,
                "_Original Abbreviation": auto_abbrev,
                "_Age": age,
                "_Color": display_hex_color(division.get("ColorHex")),
            }
        )

    return rows


def resolve_division_settings(edited_rows: list[dict], original_rows: list[dict]) -> dict:
    """
    If one abbreviation changes for a level group, apply that same suffix
    to every division with that level group while keeping each age.

    Example:
    11 Classic changed from 11C to 11CL
    12 Classic becomes 12CL
    13 Classic becomes 13CL

    Colors are not displayed or edited in the app. They stay from AES.
    """
    original_by_name = {
        row["Division Name"]: row
        for row in original_rows
    }

    edited_by_name = {
        row["Division Name"]: row
        for row in edited_rows
    }

    changed_suffix_by_level = {}

    for original in original_rows:
        division_name = original["Division Name"]
        edited = edited_by_name.get(division_name, original)

        level = original["Level Group"]
        age = original["_Age"]

        original_abbrev = normalize_power_text(str(original["_Original Abbreviation"]).strip())
        edited_abbrev = normalize_power_text(str(edited.get("Final Abbreviation", original_abbrev)).strip())

        original_suffix = abbreviation_suffix(original_abbrev, age)
        edited_suffix = abbreviation_suffix(edited_abbrev, age)

        if edited_suffix and edited_suffix != original_suffix:
            changed_suffix_by_level[level] = edited_suffix

    division_settings = {}

    for original in original_rows:
        division_id = original["DivisionId"]
        division_name = original["Division Name"]
        edited = edited_by_name.get(division_name, original)

        level = original["Level Group"]
        age = original["_Age"]

        edited_abbrev = normalize_power_text(
            str(edited.get("Final Abbreviation", original["_Original Abbreviation"])).strip()
        )

        if level in changed_suffix_by_level:
            suffix = changed_suffix_by_level[level]
            final_abbrev = f"{age}{suffix}" if age else suffix
        else:
            final_abbrev = edited_abbrev

        division_settings[division_id] = {
            "abbreviation": normalize_power_text(final_abbrev),
            "color": original["_Color"],
        }

    return division_settings


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

        match_cell.value = build_assignment_match_display_name(match, division_settings)
        match_cell.font = Font(name="Calibri", size=11, bold=True)
        match_cell.alignment = Alignment(horizontal="center", vertical="center")

        correction = get_division_correction(match, division_settings)
        fill_color = normalize_hex_color(correction["color"])
        match_cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

        format_cell.value = determine_match_format(match)
        format_cell.font = Font(name="Calibri", size=11, bold=False)
        format_cell.alignment = Alignment(horizontal="center", vertical="center")

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

        cell.value = build_worksheet_match_display_name(match, division_settings)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        correction = get_division_correction(match, division_settings)
        fill_color = normalize_hex_color(correction["color"])
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    ws.column_dimensions["A"].width = 10

    for col in range(start_court_col, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15

    for r in range(1, max_row + 1):
        ws.row_dimensions[r].height = 18

    ws.freeze_panes = "B2"


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
    "adjust division abbreviations if needed, then generate both Excel workbooks."
)

if "event_info" not in st.session_state:
    st.session_state.event_info = None

if "event_key" not in st.session_state:
    st.session_state.event_key = ""

if "generated_workbooks" not in st.session_state:
    st.session_state.generated_workbooks = None

if "division_rows" not in st.session_state:
    st.session_state.division_rows = None


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
                st.session_state.division_rows = build_division_rows(
                    st.session_state.event_info.get("Divisions", [])
                )

            st.success("Loaded.")
        except Exception:
            st.session_state.event_info = None
            st.session_state.generated_workbooks = None
            st.error("Could not load event.")


event_info = st.session_state.event_info

if event_info:
    event_name = event_info.get("Name", "")
    location = event_info.get("Location", "")
    start_date = parse_event_date(event_info["StartDate"])
    end_date = parse_event_date(event_info["EndDate"])

    st.subheader("Event Information")
    st.markdown(f"### {event_name}")

    c1, c2, c3 = st.columns([1.3, 1.3, 2.8])

    c1.metric("Start Date", format_long_date(start_date))
    c2.metric("End Date", format_long_date(end_date))
    c3.metric("Location", location or "Not listed")

    st.caption(f"AES event key: `{st.session_state.event_key}`")

    st.subheader("Division Abbreviations")

    st.write(
        "Edit abbreviations if needed. If you change one abbreviation for a shared level group, "
        "such as Classic, the same suffix will apply to all Classic divisions regardless of age."
    )

    if st.session_state.division_rows is None:
        st.session_state.division_rows = build_division_rows(event_info.get("Divisions", []))

    edited_rows = st.data_editor(
        st.session_state.division_rows,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        column_order=[
            "Division Name",
            "Level Group",
            "Team Count",
            "AES Code",
            "Final Abbreviation",
        ],
        column_config={
            "Division Name": st.column_config.TextColumn(
                "Division Name",
                disabled=True,
                width="large",
            ),
            "Level Group": st.column_config.TextColumn(
                "Level Group",
                disabled=True,
                width="medium",
            ),
            "Team Count": st.column_config.NumberColumn(
                "Team Count",
                disabled=True,
                width="small",
            ),
            "AES Code": st.column_config.TextColumn(
                "AES Code",
                disabled=True,
                width="small",
            ),
            "Final Abbreviation": st.column_config.TextColumn(
                "Final Abbreviation",
                help="Example: 18O, 17P, 14C. Changing Classic from 11C to 11CL will make 12CL, 13CL, etc.",
                width="medium",
            ),
        },
        key="division_editor",
    )

    division_settings = resolve_division_settings(
        edited_rows=edited_rows,
        original_rows=st.session_state.division_rows,
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

else:
    st.info("Enter an AES event key or full AES schedule URL to begin.")