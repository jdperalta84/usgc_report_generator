import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
import io

st.set_page_config(page_title="USGC Open CAPA Report Generator", layout="wide")
st.title("USGC Open CAPA / PAR / PTO Report Generator")
st.caption("Upload one or both annual tracking files to generate the report.")

# USGC location mapping: Location ID -> tab name (matching existing report)
USGC_LOCATIONS = {
    "120":  "HTC - 120",
    "120W": "HTC - 120W",
    "120a": "Mont Belvieu - 120a",
    "124":  "Bostco -124",
    "128":  "HOFTI - 128",
    "135":  "Minot ND - 135",
    "150":  "Texas City - 150",
    "155":  "Corpus - 155",
    "155A": "Ingelside - 155a",
    "157":  "Freeport 157",
    "156":  "Pecos - 156",
    "158":  "Brownsville - 158",
    "159":  "Cushing OK - 159",
    "160":  "Port Aurthur - 160",
    "160A": "Port Aurthur Blend - 160a",
    "161":  "Baytown TX - 161",
    "162":  "Port Lavaca 162",
    "163":  "Midland TX 163",
    "165":  "Lake Charles - 165",
}

# ── Styles ──────────────────────────────────────────────────────────────────

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def make_header_style(hex_color):
    fill = PatternFill("solid", start_color=hex_color, end_color=hex_color)
    font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return fill, font, align

def data_font():
    return Font(name="Arial", size=9)

def apply_header_row(ws, row_num, headers, hex_color):
    fill, font, align = make_header_style(hex_color)
    for col, val in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.fill = fill
        cell.font = font
        cell.alignment = align
        cell.border = thin_border()

def apply_data_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.font = data_font()
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    cell.border = thin_border()

# ── Data Loading ─────────────────────────────────────────────────────────────

def normalize_loc_id(raw):
    if raw is None:
        return ""
    if isinstance(raw, (int, float)):
        return str(int(raw))
    return str(raw).strip()

def load_cars(wb, year_suffix):
    """Load open CARs. Col M (index 12) = approved/closed date."""
    ws = wb[f"CAR '{year_suffix}"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        loc_id = normalize_loc_id(row[1])
        if not loc_id or loc_id not in USGC_LOCATIONS:
            continue
        closed = row[12]  # Col M
        if closed:  # has a closed date → skip
            continue
        description = str(row[6]).strip() if row[6] else ""
        if not description or description.upper().startswith("VOID"):
            continue
        car_num = str(row[3]) if row[3] else ""
        car_type = str(row[4]) if row[4] else ""
        area = str(row[7]) if row[7] else ""
        date_opened = row[9]
        action_plan = str(row[10]) if row[10] else ""
        completion_date = row[11]
        submitted_corp = row[13] if len(row) > 13 else None  # col N
        notes = str(row[16]) if len(row) > 16 and row[16] else ""
        rows.append({
            "loc_id": loc_id,
            "car_num": car_num,
            "car_type": car_type,
            "description": description,
            "area": area,
            "date_opened": date_opened,
            "action_plan": action_plan,
            "completion_date": completion_date,
            "submitted_corp": submitted_corp,
            "notes": notes,
        })
    return rows

def load_pars(wb, year_suffix):
    """Load open PARs. Col I (index 8) = date closed."""
    ws = wb[f"PAR '{year_suffix}"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        loc_id = normalize_loc_id(row[1])
        if not loc_id or loc_id not in USGC_LOCATIONS:
            continue
        closed = row[8]
        if closed:
            continue
        description = str(row[5]).strip() if row[5] else ""
        if not description or description.upper().startswith("VOID"):
            continue
        par_num = str(row[3]) if row[3] else ""
        par_type = str(row[4]) if row[4] else ""
        area = str(row[9]) if len(row) > 9 and row[9] else ""
        date_opened = row[6]
        notes = str(row[10]) if len(row) > 10 and row[10] else ""
        rows.append({
            "loc_id": loc_id,
            "par_num": par_num,
            "par_type": par_type,
            "description": description,
            "area": area,
            "date_opened": date_opened,
            "notes": notes,
        })
    return rows

def load_ptos(wb, year_suffix):
    """Load open PTOs. Col K (index 10) = effectiveness review / closed date."""
    ws = wb[f"PTO '{year_suffix}"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        loc_id = normalize_loc_id(row[1])
        if not loc_id or loc_id not in USGC_LOCATIONS:
            continue
        closed = row[10]
        if closed:
            continue
        description = str(row[4]).strip() if row[4] else ""
        if not description or description.upper().startswith("VOID"):
            continue
        pto_num = str(row[3]) if row[3] else ""
        pt_program = str(row[5]) if row[5] else ""
        parameter = str(row[6]) if row[6] else ""
        z_score = row[7] if row[7] else ""
        date_opened = row[8]
        notes = str(row[11]) if len(row) > 11 and row[11] else ""
        rows.append({
            "loc_id": loc_id,
            "pto_num": pto_num,
            "description": description,
            "pt_program": pt_program,
            "parameter": parameter,
            "z_score": z_score,
            "date_opened": date_opened,
            "notes": notes,
        })
    return rows

# ── Report Builder ────────────────────────────────────────────────────────────

CAR_HEADERS = ["CAR #", "Type", "Description", "Area", "Date Opened", "Days Open", "Action Plan by Due?", "Completion Date", "Notes"]
PAR_HEADERS = ["PAR #", "Type", "Description", "Area", "Date Opened", "Days Open", "Notes"]
PTO_HEADERS = ["PTO #", "Description", "PT Program", "Parameter(s)", "Z-Score", "Date Opened", "Days Open", "Notes"]

CAR_WIDTHS  = [18, 18, 50, 12, 14, 12, 20, 16, 40]
PAR_WIDTHS  = [18, 18, 60, 12, 14, 12, 40]
PTO_WIDTHS  = [18, 55, 20, 30, 10, 14, 12, 40]

CAR_COLOR = "1F4E79"   # dark blue
PAR_COLOR = "375623"   # dark green
PTO_COLOR = "7B3F00"   # dark brown

TODAY = date.today()

def days_open(d):
    if d and hasattr(d, "date"):
        return (TODAY - d.date()).days
    elif d and isinstance(d, date):
        return (TODAY - d).days
    return ""

def fmt_date(d):
    if d and hasattr(d, "strftime"):
        return d.strftime("%m/%d/%Y")
    return ""

def write_section_title(ws, row_num, title, hex_color, col_count):
    cell = ws.cell(row=row_num, column=1, value=title)
    cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    cell.fill = PatternFill("solid", start_color=hex_color, end_color=hex_color)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=col_count)
    ws.row_dimensions[row_num].height = 18

def build_location_sheet(ws, loc_id, cars, pars, ptos):
    ws.sheet_view.showGridLines = True
    current_row = 1

    # ── CARs ──────────────────────────────────────────────────────────────
    write_section_title(ws, current_row, "CORRECTIVE ACTION REQUESTS (CARs) — OPEN", CAR_COLOR, len(CAR_HEADERS))
    current_row += 1
    apply_header_row(ws, current_row, CAR_HEADERS, CAR_COLOR)
    current_row += 1

    loc_cars = [c for c in cars if c["loc_id"] == loc_id]
    if loc_cars:
        for c in loc_cars:
            vals = [
                c["car_num"],
                c["car_type"],
                c["description"],
                c["area"],
                fmt_date(c["date_opened"]),
                days_open(c["date_opened"]),
                c["action_plan"],
                fmt_date(c["completion_date"]),
                c["notes"],
            ]
            for col, v in enumerate(vals, 1):
                apply_data_cell(ws.cell(row=current_row, column=col), v)
            ws.row_dimensions[current_row].height = 40
            current_row += 1
    else:
        cell = ws.cell(row=current_row, column=1, value="No open CARs")
        cell.font = Font(italic=True, color="808080", name="Arial", size=9)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(CAR_HEADERS))
        current_row += 1

    current_row += 1  # blank row

    # ── PARs ──────────────────────────────────────────────────────────────
    write_section_title(ws, current_row, "PREVENTIVE ACTION REQUESTS (PARs) — OPEN", PAR_COLOR, len(PAR_HEADERS))
    current_row += 1
    apply_header_row(ws, current_row, PAR_HEADERS, PAR_COLOR)
    current_row += 1

    loc_pars = [p for p in pars if p["loc_id"] == loc_id]
    if loc_pars:
        for p in loc_pars:
            vals = [
                p["par_num"],
                p["par_type"],
                p["description"],
                p["area"],
                fmt_date(p["date_opened"]),
                days_open(p["date_opened"]),
                p["notes"],
            ]
            for col, v in enumerate(vals, 1):
                apply_data_cell(ws.cell(row=current_row, column=col), v)
            ws.row_dimensions[current_row].height = 40
            current_row += 1
    else:
        cell = ws.cell(row=current_row, column=1, value="No open PARs")
        cell.font = Font(italic=True, color="808080", name="Arial", size=9)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(PAR_HEADERS))
        current_row += 1

    current_row += 1

    # ── PTOs ──────────────────────────────────────────────────────────────
    write_section_title(ws, current_row, "PROFICIENCY TESTING OUTLIERS (PTOs) — OPEN", PTO_COLOR, len(PTO_HEADERS))
    current_row += 1
    apply_header_row(ws, current_row, PTO_HEADERS, PTO_COLOR)
    current_row += 1

    loc_ptos = [p for p in ptos if p["loc_id"] == loc_id]
    if loc_ptos:
        for p in loc_ptos:
            vals = [
                p["pto_num"],
                p["description"],
                p["pt_program"],
                p["parameter"],
                p["z_score"],
                fmt_date(p["date_opened"]),
                days_open(p["date_opened"]),
                p["notes"],
            ]
            for col, v in enumerate(vals, 1):
                apply_data_cell(ws.cell(row=current_row, column=col), v)
            ws.row_dimensions[current_row].height = 40
            current_row += 1
    else:
        cell = ws.cell(row=current_row, column=1, value="No open PTOs")
        cell.font = Font(italic=True, color="808080", name="Arial", size=9)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(PTO_HEADERS))
        current_row += 1

    # Column widths (use max of CAR widths since it has most columns)
    max_widths = {}
    for i, w in enumerate(CAR_WIDTHS, 1):
        max_widths[i] = w
    for i, w in enumerate(PAR_WIDTHS, 1):
        max_widths[i] = max(max_widths.get(i, 0), w)
    for i, w in enumerate(PTO_WIDTHS, 1):
        max_widths[i] = max(max_widths.get(i, 0), w)
    for col_idx, width in max_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def build_cover_sheet(ws, cars, pars, ptos):
    today_str = TODAY.strftime("%B %d, %Y")

    # Title
    ws.merge_cells("A1:L1")
    title_cell = ws["A1"]
    title_cell.value = f"USGC Open CAPA / PAR / PTO Report — As of {today_str}"
    title_cell.font = Font(bold=True, name="Arial", size=13, color="FFFFFF")
    title_cell.fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Header row
    cover_headers = [
        "Location", "Loc ID",
        "CARs +90d", "CARs +30d", "CARs -30d", "CARs Total",
        "PARs Open",
        "PTOs Open",
    ]
    hdr_colors = ["1F4E79"] * 2 + ["C00000"] * 4 + ["375623"] + ["7B3F00"]
    for col, (h, c) in enumerate(zip(cover_headers, hdr_colors), 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill = PatternFill("solid", start_color=c, end_color=c)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[2].height = 30

    row = 3
    for loc_id, tab_name in USGC_LOCATIONS.items():
        loc_cars = [c for c in cars if c["loc_id"] == loc_id]
        loc_pars = [p for p in pars if p["loc_id"] == loc_id]
        loc_ptos = [p for p in ptos if p["loc_id"] == loc_id]

        car_days = [days_open(c["date_opened"]) for c in loc_cars if isinstance(days_open(c["date_opened"]), int)]
        cars_90 = sum(1 for d in car_days if d > 90)
        cars_30 = sum(1 for d in car_days if 30 < d <= 90)
        cars_under30 = sum(1 for d in car_days if d <= 30)
        cars_total = len(loc_cars)

        vals = [tab_name, loc_id, cars_90, cars_30, cars_under30, cars_total, len(loc_pars), len(loc_ptos)]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=v)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="center" if col > 2 else "left", vertical="center")
            cell.border = thin_border()
            # Red highlight if any +90 day CARs
            if col == 3 and v > 0:
                cell.fill = PatternFill("solid", start_color="FFD7D7", end_color="FFD7D7")
        row += 1

    # Totals row
    ws.cell(row=row, column=1, value="TOTALS").font = Font(bold=True, name="Arial", size=10)
    for col in range(3, 9):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col, value=f"=SUM({col_letter}3:{col_letter}{row-1})")
        cell.font = Font(bold=True, name="Arial", size=10)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 10
    for col in range(3, 9):
        ws.column_dimensions[get_column_letter(col)].width = 13


def generate_report(files_data):
    all_cars, all_pars, all_ptos = [], [], []

    for file_bytes, year_suffix in files_data:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        all_cars += load_cars(wb, year_suffix)
        all_pars += load_pars(wb, year_suffix)
        all_ptos += load_ptos(wb, year_suffix)

    out_wb = openpyxl.Workbook()
    cover = out_wb.active
    cover.title = "Cover Sheet"

    for loc_id, tab_name in USGC_LOCATIONS.items():
        ws = out_wb.create_sheet(title=tab_name)
        build_location_sheet(ws, loc_id, all_cars, all_pars, all_ptos)

    build_cover_sheet(cover, all_cars, all_pars, all_ptos)

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ────────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)
with col1:
    file_2025 = st.file_uploader("2025 Tracking File", type=["xlsx"], key="f25")
with col2:
    file_2026 = st.file_uploader("2026 Tracking File", type=["xlsx"], key="f26")

if not file_2025 and not file_2026:
    st.info("Upload at least one tracking file to generate the report.")
else:
    files_data = []
    if file_2025:
        files_data.append((file_2025.read(), "25"))
    if file_2026:
        files_data.append((file_2026.read(), "26"))

    if st.button("Generate Report", type="primary"):
        with st.spinner("Building report..."):
            try:
                buf = generate_report(files_data)
                fname = f"Open_CAPA_PTO_USGC_as_of_{TODAY.strftime('%m-%d-%y')}.xlsx"
                st.success("Report generated successfully!")

                # Quick summary
                all_cars2, all_pars2, all_ptos2 = [], [], []
                for file_bytes, yr in files_data:
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                    all_cars2 += load_cars(wb, yr)
                    all_pars2 += load_pars(wb, yr)
                    all_ptos2 += load_ptos(wb, yr)

                c1, c2, c3 = st.columns(3)
                c1.metric("Open CARs (USGC)", len(all_cars2))
                c2.metric("Open PARs (USGC)", len(all_pars2))
                c3.metric("Open PTOs (USGC)", len(all_ptos2))

                st.download_button(
                    label="⬇️ Download Report",
                    data=buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error generating report: {e}")
                st.exception(e)
