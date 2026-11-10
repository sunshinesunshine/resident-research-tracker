import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette (Medical Academic Professional)
NAVY_HEADER = "1F4E79"       # Primary Header Deep Blue
NAVY_SUBHEADER = "2F5597"    # Section Banner
TEAL_ACCENT = "008080"       # Accent Teal
LIGHT_BLUE_FILL = "EDF2F8"   # Alternating row fill
ZEBRA_FILL = "F9FAFC"        # Very light row fill
KPI_BG = "F0F4F8"            # KPI card fill
SUCCESS_GREEN = "D9EAD3"     # Completed light green
SUCCESS_TEXT = "274E13"      # Completed dark green
WARNING_AMBER = "FCE5CD"     # Behind Schedule light orange
WARNING_TEXT = "783F04"      # Behind Schedule dark orange
BORDER_COLOR = "D9D9D9"      # Light grey border
WHITE = "FFFFFF"

thin_side = Side(border_style="thin", color=BORDER_COLOR)
med_side = Side(border_style="medium", color="1F4E79")
border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_header = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

font_family = "Segoe UI"
title_font = Font(name=font_family, size=16, bold=True, color="1F4E79")
subtitle_font = Font(name=font_family, size=10, italic=True, color="595959")
header_font = Font(name=font_family, size=10, bold=True, color=WHITE)
sub_header_font = Font(name=font_family, size=10, bold=True, color="1F4E79")
data_font = Font(name=font_family, size=10, color="000000")
data_bold = Font(name=font_family, size=10, bold=True, color="000000")
kpi_num_font = Font(name=font_family, size=18, bold=True, color="1F4E79")
kpi_label_font = Font(name=font_family, size=9, bold=True, color="595959")

# -------------------------------------------------------------
# TAB 1: Dashboard
# -------------------------------------------------------------
ws_dash = wb.create_sheet(title="Dashboard")
ws_dash.views.sheetView[0].showGridLines = True

# Title
ws_dash["A1"] = "ระบบติดตามความก้าวหน้างานวิจัยแพทย์ประจำบ้าน (Resident Research Progress Tracker)"
ws_dash["A1"].font = title_font
ws_dash["A2"] = "ภาควิชาจักษุวิทยา | ภาพรวม 3 ชั้นปี (R1, R2, R3) | ตรวจสอบข้อมูลแบบ Real-time"
ws_dash["A2"].font = subtitle_font

# KPI Cards row 4 to 6
kpi_configs = [
    ("B4", "C4", "B5", "C5", "จำนวน Resident ทั้งหมด", "=COUNTA(R1_2572!B4:B9)+COUNTA(R2_2571!B4:B9)+COUNTA(R3_2570!B4:B9)", "คน"),
    ("E4", "F4", "E5", "F5", "เสร็จสมบูรณ์แล้ว (100%)", '=COUNTIF(R1_2572!I4:I9, "Completed")+COUNTIF(R2_2571!I4:I9, "Completed")+COUNTIF(R3_2570!I4:I9, "Completed")', "โครงการ"),
    ("H4", "I4", "H5", "I5", "กำลังดำเนินงาน (On Track)", '=COUNTIF(R1_2572!I4:I9, "On Track")+COUNTIF(R2_2571!I4:I9, "On Track")+COUNTIF(R3_2570!I4:I9, "On Track")', "โครงการ"),
    ("K4", "L4", "K5", "L5", "ต้องติดตามเร่งด่วน (Behind Schedule)", '=COUNTIF(R1_2572!I4:I9, "Behind Schedule")+COUNTIF(R2_2571!I4:I9, "Behind Schedule")+COUNTIF(R3_2570!I4:I9, "Behind Schedule")', "โครงการ"),
]

for top_l_c, top_r_c, bot_l_c, bot_r_c, label, formula, unit in kpi_configs:
    ws_dash.merge_cells(f"{top_l_c}:{top_r_c}")
    ws_dash.merge_cells(f"{bot_l_c}:{bot_r_c}")
    cell_top = ws_dash[top_l_c]
    cell_top.value = label
    cell_top.font = kpi_label_font
    cell_top.alignment = Alignment(horizontal="center", vertical="center")
    cell_top.fill = PatternFill(start_color=KPI_BG, end_color=KPI_BG, fill_type="solid")
    
    cell_bot = ws_dash[bot_l_c]
    cell_bot.value = formula
    cell_bot.font = kpi_num_font
    cell_bot.alignment = Alignment(horizontal="center", vertical="center")
    cell_bot.fill = PatternFill(start_color=KPI_BG, end_color=KPI_BG, fill_type="solid")

    for r in range(int(top_l_c[1:]), int(bot_l_c[1:]) + 1):
        for col_name in [top_l_c[0], top_r_c[0]]:
            ws_dash[f"{col_name}{r}"].border = border_thin

# Cohort Summary Table
ws_dash["A8"] = "1. สรุปความคืบหน้ารายชั้นปี (Summary by Residency Cohort)"
ws_dash["A8"].font = Font(name=font_family, size=11, bold=True, color="1F4E79")

summary_headers = ["ชั้นปี", "รุ่นปีการศึกษา (จบปี)", "จำนวน (คน)", "ความคืบหน้าเฉลี่ย (%)", "เสร็จสิ้น (100%)", "On Track", "Behind Schedule", "ขั้นตอนหลักที่ควรอยู่ ณ ปัจจุบัน"]
for col_idx, h in enumerate(summary_headers, start=1):
    cell = ws_dash.cell(row=9, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_header

cohort_rows = [
    ("แพทย์ประจำบ้านชั้นปีที่ 1 (R1)", "พ.ศ. 2569 (จบ 2572)", "=COUNTA(R1_2572!B4:B9)", "=AVERAGE(R1_2572!H4:H9)", '=COUNTIF(R1_2572!I4:I9, "Completed")', '=COUNTIF(R1_2572!I4:I9, "On Track")', '=COUNTIF(R1_2572!I4:I9, "Behind Schedule")', "1. ทบทวนวรรณกรรม ➔ 2. ยื่นโครงร่าง IRB"),
    ("แพทย์ประจำบ้านชั้นปีที่ 2 (R2)", "พ.ศ. 2568 (จบ 2571)", "=COUNTA(R2_2571!B4:B9)", "=AVERAGE(R2_2571!H4:H9)", '=COUNTIF(R2_2571!I4:I9, "Completed")', '=COUNTIF(R2_2571!I4:I9, "On Track")', '=COUNTIF(R2_2571!I4:I9, "Behind Schedule")', "3. ขอทุน/แก้ IRB ➔ 4. เก็บข้อมูลงานวิจัย"),
    ("แพทย์ประจำบ้านชั้นปีที่ 3 (R3)", "พ.ศ. 2567 (จบ 2570)", "=COUNTA(R3_2570!B4:B9)", "=AVERAGE(R3_2570!H4:H9)", '=COUNTIF(R3_2570!I4:I9, "Completed")', '=COUNTIF(R3_2570!I4:I9, "On Track")', '=COUNTIF(R3_2570!I4:I9, "Behind Schedule")', "5. วิเคราะห์สถิติ ➔ 6. รูปเล่ม & เผยแพร่"),
]

for row_idx, data in enumerate(cohort_rows, start=10):
    for col_idx, val in enumerate(data, start=1):
        cell = ws_dash.cell(row=row_idx, column=col_idx, value=val)
        cell.font = data_font
        cell.border = border_thin
        if col_idx in [1, 2, 8]:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx == 4:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0.0%"
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")

# Total summary row
ws_dash.cell(row=13, column=1, value="รวมทั้งหมด (Total)").font = data_bold
ws_dash.cell(row=13, column=1).alignment = Alignment(horizontal="center", vertical="center")
ws_dash.cell(row=13, column=2, value="-").alignment = Alignment(horizontal="center", vertical="center")
ws_dash.cell(row=13, column=3, value="=SUM(C10:C12)").number_format = "#,##0"
ws_dash.cell(row=13, column=4, value="=AVERAGE(D10:D12)").number_format = "0.0%"
ws_dash.cell(row=13, column=5, value="=SUM(E10:E12)").number_format = "#,##0"
ws_dash.cell(row=13, column=6, value="=SUM(F10:F12)").number_format = "#,##0"
ws_dash.cell(row=13, column=7, value="=SUM(G10:G12)").number_format = "#,##0"
ws_dash.cell(row=13, column=8, value="ทุกชั้นปีอยู่ในการติดตาม").alignment = Alignment(horizontal="center", vertical="center")

for c in range(1, 9):
    cell = ws_dash.cell(row=13, column=c)
    cell.font = data_bold
    cell.fill = PatternFill(start_color=LIGHT_BLUE_FILL, end_color=LIGHT_BLUE_FILL, fill_type="solid")
    cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=Side(border_style="double", color="1F4E79"))
    if c in [3, 4, 5, 6, 7]:
        cell.alignment = Alignment(horizontal="center", vertical="center")

# Table 2: Action Items / Behind schedule monitoring
ws_dash["A16"] = "2. รายชื่อแพทย์ประจำบ้านที่ต้องติดตามเป็นพิเศษ (Action Required / Behind Schedule Alerts)"
ws_dash["A16"].font = Font(name=font_family, size=11, bold=True, color="C00000")

action_headers = ["ชั้นปี", "ชื่อ-นามสกุล", "หัวข้องานวิจัย", "ประเภทงานวิจัย", "อาจารย์ที่ปรึกษา", "ความคืบหน้า (%)", "สถานะ", "ขั้นตอนที่ยังค้างดำเนินการ"]
for col_idx, h in enumerate(action_headers, start=1):
    cell = ws_dash.cell(row=17, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = PatternFill(start_color="843C0C", end_color="843C0C", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_header

# Sample flagged residents linked directly to the sheets
sample_action_rows = [
    ("R2", "=R2_2571!B6", "=R2_2571!D6", "=R2_2571!E6", "=R2_2571!F6", "=R2_2571!H6", "=R2_2571!I6", "ยังไม่ผ่านขั้นตอน 4: เก็บข้อมูลงานวิจัย"),
    ("R3", "=R3_2570!B7", "=R3_2570!D7", "=R3_2570!E7", "=R3_2570!F7", "=R3_2570!H7", "=R3_2570!I7", "ยังไม่ผ่านขั้นตอน 5: วิเคราะห์ทางสถิติ"),
]

for row_idx, data in enumerate(sample_action_rows, start=18):
    for col_idx, val in enumerate(data, start=1):
        cell = ws_dash.cell(row=row_idx, column=col_idx, value=val)
        cell.font = data_font
        cell.border = border_thin
        if col_idx in [1, 7]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx == 6:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0.0%"
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        if col_idx == 7:
            cell.fill = PatternFill(start_color=WARNING_AMBER, end_color=WARNING_AMBER, fill_type="solid")
            cell.font = Font(name=font_family, size=10, bold=True, color=WARNING_TEXT)

# Table 3: Research Design Breakdown in Department
ws_dash["A22"] = "3. สัดส่วนประเภทงานวิจัยของแพทย์ประจำบ้าน (Research Study Design Breakdown)"
ws_dash["A22"].font = Font(name=font_family, size=11, bold=True, color="1F4E79")

design_headers = ["ประเภทงานวิจัย (Study Design)", "R1 (คน)", "R2 (คน)", "R3 (คน)", "รวมทั้งหมด (โครงการ)"]
for col_idx, h in enumerate(design_headers, start=1):
    cell = ws_dash.cell(row=23, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_header

designs = [
    "Prospective Cohort Study",
    "Retrospective Chart Review",
    "Cross-sectional Study",
    "Randomized Controlled Trial (RCT)",
    "Diagnostic Accuracy Study",
    "Systematic Review & Meta-analysis"
]

for idx, d in enumerate(designs, start=24):
    ws_dash.cell(row=idx, column=1, value=d).alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.cell(row=idx, column=2, value=f'=COUNTIF(R1_2572!$E$4:$E$9, "{d}")').alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.cell(row=idx, column=3, value=f'=COUNTIF(R2_2571!$E$4:$E$9, "{d}")').alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.cell(row=idx, column=4, value=f'=COUNTIF(R3_2570!$E$4:$E$9, "{d}")').alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.cell(row=idx, column=5, value=f'=SUM(B{idx}:D{idx})').alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 6):
        ws_dash.cell(row=idx, column=c).font = data_font
        ws_dash.cell(row=idx, column=c).border = border_thin

# -------------------------------------------------------------
# Function to generate Cohort Sheets
# -------------------------------------------------------------
def build_cohort_sheet(sheet_title, cohort_name, entering_year, grad_year, timeline_info, sample_residents):
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True

    # Header Title Banner
    ws["A1"] = f"ตารางติดตามความก้าวหน้างงานวิจัยแพทย์ประจำบ้าน: {cohort_name}"
    ws["A1"].font = title_font
    ws["A2"] = f"รุ่นปีการศึกษา {entering_year} (สำเร็จการศึกษา พ.ศ. {grad_year}) | แผนแม่บท 6 ขั้นตอนตามเกณฑ์ภาควิชา"
    ws["A2"].font = subtitle_font

    # Column Headers (Row 3)
    headers = [
        ("A3", "ลำดับ\nNo.", 6),
        ("B3", "ชื่อ-นามสกุล\nResident Name", 24),
        ("C3", "สาขา\nSubspecialty", 14),
        ("D3", "หัวข้องานวิจัย\nResearch Title", 38),
        ("E3", "ประเภทงานวิจัย\nStudy Design", 24),
        ("F3", "อาจารย์ที่ปรึกษาหลัก\nPrimary Advisor", 22),
        ("G3", "อาจารย์ที่ปรึกษาร่วม\nCo-Advisor", 20),
        ("H3", "ความคืบหน้า\nProgress (%)", 14),
        ("I3", "สถานะ\nStatus", 16),
        ("J3", f"1. ทบทวนวรรณกรรม\n({timeline_info[0]})", 18),
        ("K3", f"2. ยื่นโครงร่าง IRB\n({timeline_info[1]})", 18),
        ("L3", f"3. ขอทุน & แก้ไข\n({timeline_info[2]})", 18),
        ("M3", f"4. เก็บข้อมูล\n({timeline_info[3]})", 18),
        ("N3", f"5. วิเคราะห์สถิติ\n({timeline_info[4]})", 18),
        ("O3", f"6. รูปเล่ม & เผยแพร่\n({timeline_info[5]})", 18),
        ("P3", "หมายเหตุ / ลิงก์ Drive\nRemarks & Drive Link", 28),
    ]

    for cell_id, text, width in headers:
        cell = ws[cell_id]
        cell.value = text
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_header
        
        # Color coding for header sections
        col_letter = cell_id[0] if len(cell_id) == 2 else cell_id[:2]
        if col_letter in ["A", "B", "C", "D", "E", "F", "G"]:
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        elif col_letter in ["H", "I"]:
            cell.fill = PatternFill(start_color=TEAL_ACCENT, end_color=TEAL_ACCENT, fill_type="solid")
        elif col_letter in ["J", "K", "L", "M", "N", "O"]:
            cell.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
        
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[3].height = 42

    # Insert Rows
    for idx, r in enumerate(sample_residents, start=4):
        ws.row_dimensions[idx].height = 24
        # A: No
        c_a = ws.cell(row=idx, column=1, value=idx-3)
        c_a.alignment = Alignment(horizontal="center", vertical="center")
        
        # B: Name
        c_b = ws.cell(row=idx, column=2, value=r["name"])
        c_b.alignment = Alignment(horizontal="left", vertical="center")
        
        # C: Subspecialty
        c_c = ws.cell(row=idx, column=3, value=r["subspecialty"])
        c_c.alignment = Alignment(horizontal="center", vertical="center")
        
        # D: Title
        c_d = ws.cell(row=idx, column=4, value=r["title"])
        c_d.alignment = Alignment(horizontal="left", vertical="center")
        
        # E: Study Design
        c_e = ws.cell(row=idx, column=5, value=r["design"])
        c_e.alignment = Alignment(horizontal="left", vertical="center")
        
        # F: Primary Advisor
        c_f = ws.cell(row=idx, column=6, value=r["advisor"])
        c_f.alignment = Alignment(horizontal="left", vertical="center")
        
        # G: Co-Advisor
        c_g = ws.cell(row=idx, column=7, value=r["co_advisor"])
        c_g.alignment = Alignment(horizontal="left", vertical="center")

        # Checkboxes (J to O: Cols 10 to 15)
        for c_idx, check_val in enumerate(r["checks"], start=10):
            c_check = ws.cell(row=idx, column=c_idx, value=check_val)
            c_check.alignment = Alignment(horizontal="center", vertical="center")
            c_check.font = data_bold
            # Light green tint if TRUE
            if check_val:
                c_check.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                c_check.font = Font(name=font_family, size=10, bold=True, color="274E13")

        # H: Progress Formula =COUNTIF(J4:O4, TRUE)/6
        c_h = ws.cell(row=idx, column=8, value=f"=COUNTIF(J{idx}:O{idx}, TRUE)/6")
        c_h.alignment = Alignment(horizontal="center", vertical="center")
        c_h.number_format = "0.0%"
        c_h.font = data_bold

        # I: Status Formula
        status_formula = r["status_formula"].replace("{row}", str(idx))
        c_i = ws.cell(row=idx, column=9, value=status_formula)
        c_i.alignment = Alignment(horizontal="center", vertical="center")
        c_i.font = data_bold
        if r.get("flag") == "Behind Schedule":
            c_i.fill = PatternFill(start_color=WARNING_AMBER, end_color=WARNING_AMBER, fill_type="solid")
            c_i.font = Font(name=font_family, size=10, bold=True, color=WARNING_TEXT)
        elif r.get("flag") == "Completed":
            c_i.fill = PatternFill(start_color=SUCCESS_GREEN, end_color=SUCCESS_GREEN, fill_type="solid")
            c_i.font = Font(name=font_family, size=10, bold=True, color=SUCCESS_TEXT)
        else:
            c_i.fill = PatternFill(start_color=LIGHT_BLUE_FILL, end_color=LIGHT_BLUE_FILL, fill_type="solid")
            c_i.font = Font(name=font_family, size=10, bold=True, color="1F4E79")

        # P: Remarks / Drive Link
        c_p = ws.cell(row=idx, column=16, value=r.get("remarks", "https://drive.google.com/..."))
        c_p.alignment = Alignment(horizontal="left", vertical="center")
        c_p.font = Font(name=font_family, size=9, italic=True, color="0000FF")

        # Apply borders and basic font
        for c in range(1, 17):
            cell = ws.cell(row=idx, column=c)
            cell.border = border_thin
            if cell.font == data_font or cell.font is None:
                cell.font = data_font

    # Summary Row at the bottom
    tot_row = len(sample_residents) + 4
    ws.cell(row=tot_row, column=1, value="เฉลี่ยรวม").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tot_row, column=1).font = data_bold
    ws.cell(row=tot_row, column=2, value=f"=COUNTA(B4:B{tot_row-1})&\" คน\"").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tot_row, column=2).font = data_bold
    ws.cell(row=tot_row, column=8, value=f"=AVERAGE(H4:H{tot_row-1})").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tot_row, column=8).font = data_bold
    ws.cell(row=tot_row, column=8).number_format = "0.0%"

    for c_idx in range(10, 16):
        col_let = get_column_letter(c_idx)
        cell = ws.cell(row=tot_row, column=c_idx, value=f'=COUNTIF({col_let}4:{col_let}{tot_row-1}, TRUE)&"/{len(sample_residents)}"')
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = data_bold

    for c in range(1, 17):
        cell = ws.cell(row=tot_row, column=c)
        cell.fill = PatternFill(start_color=LIGHT_BLUE_FILL, end_color=LIGHT_BLUE_FILL, fill_type="solid")
        cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=Side(border_style="double", color="1F4E79"))

    # Instruction Guide beneath table
    guide_row = tot_row + 2
    ws.cell(row=guide_row, column=2, value="💡 วิธีการอัปเดตสำหรับ Resident:").font = Font(name=font_family, size=10, bold=True, color="1F4E79")
    ws.cell(row=guide_row+1, column=2, value="1. เมื่อผ่านแต่ละขั้นตอน ให้คลิกใส่เครื่องหมายถูก (TRUE) ในช่อง Checkbox คอลัมน์ J - O").font = subtitle_font
    ws.cell(row=guide_row+2, column=2, value="2. ระบบจะคำนวณ % ความคืบหน้า (Progress) และปรับสถานะ On Track / Completed / Behind Schedule ให้โดยอัตโนมัติ").font = subtitle_font
    ws.cell(row=guide_row+3, column=2, value="3. สามารถวางลิงก์ Google Drive รวบรวมเอกสาร Proposal, IRB Certificate, หรือ Dataset ในคอลัมน์ P เพื่อให้อาจารย์ที่ปรึกษาเข้าตรวจได้ทันที").font = subtitle_font

# -------------------------------------------------------------
# Data Definition for R1, R2, R3
# -------------------------------------------------------------

# R1 (จบปี 2572)
# Timeline: ก.ค. 69 - ต.ค. 69, พ.ย. 69 - ก.พ. 70, มี.ค. 70 - มิ.ย. 70, ก.ค. 70 - ต.ค. 70, พ.ย. 70 - ม.ค. 71, ก.พ. 71 - มิ.ย. 72
r1_timeline = ["ก.ค. 69 - ต.ค. 69", "พ.ย. 69 - ก.พ. 70", "มี.ค. 70 - มิ.ย. 70", "ก.ค. 70 - ต.ค. 70", "พ.ย. 70 - ม.ค. 71", "ก.พ. 71 - มิ.ย. 72"]
r1_residents = [
    {
        "name": "พญ. นภัสสร วงศ์วิวัฒน์",
        "subspecialty": "Glaucoma",
        "title": "การประเมิน Corneal Hysteresis ในผู้ป่วย Normal Tension Glaucoma",
        "design": "Prospective Cohort Study",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "รศ.นพ. สุรชัย เมตตา",
        "checks": [True, False, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ทบทวนวรรณกรรมและร่าง PICO เสร็จแล้ว เตรียมยื่น IRB"
    },
    {
        "name": "นพ. กิตติศักดิ์ เจริญผล",
        "subspecialty": "Retina",
        "title": "ความสัมพันธ์ของ OCT-Angiography กับ Diabetic Retinopathy Progression",
        "design": "Retrospective Chart Review",
        "advisor": "ผศ.นพ. วีระศักดิ์ แสงธรรม",
        "co_advisor": "-",
        "checks": [True, True, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ยื่นเอกสาร IRB รอบแรกแล้ว รอผลพิจารณา"
    },
    {
        "name": "พญ. ธนพร รัตนศิริ",
        "subspecialty": "Cornea",
        "title": "ประสิทธิภาพของน้ำตาเทียมสูตรใหม่ในการรักษา Dry Eye Disease",
        "design": "Randomized Controlled Trial (RCT)",
        "advisor": "รศ.พญ. ศิริพร ทิพยรัตน์",
        "co_advisor": "อ.นพ. ธนกฤต ชัยชนะ",
        "checks": [True, False, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "กำลังสรุป Literature Review ร่วมกับ Advisor"
    },
    {
        "name": "นพ. ชนกันต์ สุวรรณเวช",
        "subspecialty": "Pediatric",
        "title": "ปัจจัยเสี่ยงของการเกิด Myopia Progression ในเด็กวัยเรียนช่วงประถมศึกษา",
        "design": "Prospective Cohort Study",
        "advisor": "ผศ.พญ. ปัทมา วิเศษโภคัย",
        "co_advisor": "-",
        "checks": [False, False, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "Behind Schedule",
        "remarks": "ยังปรับแก้หัวข้อวิจัยกับอาจารย์ที่ปรึกษา ล่าช้ากว่ากำหนด"
    },
    {
        "name": "พญ. วริศรา เกษมสุข",
        "subspecialty": "Neuro-Oph",
        "title": "ลักษณะทางคลินิกและพยากรณ์โรคของ NAION ในประชากรไทย",
        "design": "Retrospective Chart Review",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "รศ.นพ. สุรชัย เมตตา",
        "checks": [True, False, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ทบทวนวรรณกรรมเรียบร้อย ร่างแบบเสนอโครงการเสร็จ 80%"
    },
    {
        "name": "นพ. ภูริชญ์ กุลวัฒน์",
        "subspecialty": "Oculoplastic",
        "title": "ผลการผ่าตัด External DCR เปรียบเทียบกับ Endoscopic DCR ในผู้ใหญ่",
        "design": "Comparative Cohort Study",
        "advisor": "รศ.นพ. ประพันธ์ ฤทธิ์เดช",
        "co_advisor": "-",
        "checks": [True, True, False, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(J{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ยื่นขอพิจารณาจริยธรรมแล้ว (รอรับรอง)"
    },
]

# R2 (จบปี 2571)
# Timeline: ก.ค. 68 - ต.ค. 68, พ.ย. 68 - ก.พ. 69, มี.ค. 69 - มิ.ย. 69, ก.ค. 69 - ต.ค. 69, พ.ย. 69 - ม.ค. 70, ก.พ. 70 - มิ.ย. 71
r2_timeline = ["ก.ค. 68 - ต.ค. 68", "พ.ย. 68 - ก.พ. 69", "มี.ค. 69 - มิ.ย. 69", "ก.ค. 69 - ต.ค. 69", "พ.ย. 69 - ม.ค. 70", "ก.พ. 70 - มิ.ย. 71"]
r2_residents = [
    {
        "name": "นพ. อัครเดช มงคลเจริญ",
        "subspecialty": "Glaucoma",
        "title": "การเปรียบเทียบผลลัพธ์ระยะยาวระหว่าง SubCyclo กับ cwTSCPC ในต้อหินดื้อยา",
        "design": "Prospective Comparative Study",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "รศ.นพ. สุรชัย เมตตา",
        "checks": [True, True, True, True, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "เก็บข้อมูลครบ 100 ราย กำลังเริ่มคีย์ข้อมูลลง CRF"
    },
    {
        "name": "พญ. กมลวรรณ ชัยเจริญ",
        "subspecialty": "Retina",
        "title": "ความแม่นยำของ AI ในการตรวจคัดกรองเบาหวานขึ้นจอประสาทตาผ่าน Tele-ophthalmology",
        "design": "Diagnostic Accuracy Study",
        "advisor": "ผศ.นพ. วีระศักดิ์ แสงธรรม",
        "co_advisor": "-",
        "checks": [True, True, True, True, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "เก็บชุดภาพและประเมินผล Sensitivity/Specificity ตามแผน"
    },
    {
        "name": "นพ. ณัฐวุฒิ สมบูรณ์ทรัพย์",
        "subspecialty": "Cornea",
        "title": "อุบัติการณ์และการจัดการภาวะ Graft Rejection หลังการผ่าตัด DSEK",
        "design": "Retrospective Chart Review",
        "advisor": "รศ.พญ. ศิริพร ทิพยรัตน์",
        "co_advisor": "-",
        "checks": [True, True, True, False, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "Behind Schedule",
        "remarks": "เวชระเบียนโรงพยาบาลค้นประวัติได้ล่าช้า ยังเก็บข้อมูลไม่ครบ"
    },
    {
        "name": "พญ. พิชชาภา เลิศวิไล",
        "subspecialty": "Uveitis",
        "title": "รูปแบบทางคลินิกและการตอบสนองต่อยาชีววัตถุใน Non-infectious Uveitis",
        "design": "Prospective Cohort Study",
        "advisor": "ผศ.พญ. สุดารัตน์ มีสุข",
        "co_advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "checks": [True, True, True, True, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "เก็บข้อมูลคนไข้กลุ่มเป้าหมายได้ตามกำหนดเวลา"
    },
    {
        "name": "นพ. ภัทรพล วาณิชสกุล",
        "subspecialty": "Cataract",
        "title": "การเปรียบเทียบความแม่นยำของสูตร IOL Calculation รุ่นใหม่ในตาสั้นมาก",
        "design": "Cross-sectional Study",
        "advisor": "รศ.นพ. ประพันธ์ ฤทธิ์เดช",
        "co_advisor": "-",
        "checks": [True, True, True, True, True, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "เก็บข้อมูลเสร็จแล้ว กำลังวิเคราะห์สถิติก่อนกำหนด"
    },
    {
        "name": "พญ. รินรดา พงษ์ไพศาล",
        "subspecialty": "Glaucoma",
        "title": "ผลการรักษาด้วย Micropulse Transscleral Cyclophotocoagulation ในเด็ก",
        "design": "Retrospective Case Series",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "-",
        "checks": [True, True, True, True, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(M{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ข้อมูลเก็บเรียบร้อย นัดพบนักสถิติสัปดาห์หน้า"
    },
]

# R3 (จบปี 2570)
# Timeline: ก.ค. 67 - ต.ค. 67, พ.ย. 67 - ก.พ. 68, มี.ค. 68 - มิ.ย. 68, ก.ค. 68 - ต.ค. 68, พ.ย. 68 - ม.ค. 69, ก.พ. 69 - มิ.ย. 70
r3_timeline = ["ก.ค. 67 - ต.ค. 67", "พ.ย. 67 - ก.พ. 68", "มี.ค. 68 - มิ.ย. 68", "ก.ค. 68 - ต.ค. 68", "พ.ย. 68 - ม.ค. 69", "ก.พ. 69 - มิ.ย. 70"]
r3_residents = [
    {
        "name": "นพ. ปริญญา บุญญานุภาพ",
        "subspecialty": "Glaucoma",
        "title": "ประสิทธิผลและความปลอดภัยของ MicroPulse TSCPC ในภาวะต้อหินมุมปิดเรื้อรัง",
        "design": "Prospective Cohort Study",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "รศ.นพ. สุรชัย เมตตา",
        "checks": [True, True, True, True, True, True],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "Completed",
        "remarks": "ส่งบทความตีพิมพ์ในวารสาร TJO เรียบร้อย และส่งเล่มวิจัยฉบับสมบูรณ์แล้ว"
    },
    {
        "name": "พญ. ชนากานต์ สุทธิวงศ์",
        "subspecialty": "Retina",
        "title": "การประเมิน Macular Perfusion หลังฉีด Anti-VEGF ใน Retinal Vein Occlusion",
        "design": "Prospective Cohort Study",
        "advisor": "ผศ.นพ. วีระศักดิ์ แสงธรรม",
        "co_advisor": "-",
        "checks": [True, True, True, True, True, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "วิเคราะห์สถิติเสร็จสิ้น อยู่ระหว่างเขียน Discussion และจัดรูปเล่มวิจัย"
    },
    {
        "name": "นพ. วรปรัชญ์ ศรีสวัสดิ์",
        "subspecialty": "Cornea",
        "title": "การประเมินผลทางชีวกลศาสตร์ของกระจกตาหลังทำ SMILE vs FS-LASIK",
        "design": "Prospective Comparative Study",
        "advisor": "รศ.พญ. ศิริพร ทิพยรัตน์",
        "co_advisor": "-",
        "checks": [True, True, True, True, True, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "ร่าง Manuscript ภาษาอังกฤษพร้อมส่ง Subspecialty Meeting"
    },
    {
        "name": "พญ. มนัสนันท์ ภักดีชนม์",
        "subspecialty": "Pediatric",
        "title": "ผลการรักษา Amblyopia ด้วย Optical Pharmacological Penalization",
        "design": "Randomized Controlled Trial (RCT)",
        "advisor": "ผศ.พญ. ปัทมา วิเศษโภคัย",
        "co_advisor": "-",
        "checks": [True, True, True, True, False, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "Behind Schedule",
        "remarks": "ยังรันสถิติโมเดลเชิงซ้อนไม่ผ่าน อยู่ระหว่างปรึกษาหน่วยชีวสถิติ"
    },
    {
        "name": "นพ. ธัชพล เตชะวณิช",
        "subspecialty": "Neuro-Oph",
        "title": "Optical Coherence Tomography Biomarkers in Multiple Sclerosis without Optic Neuritis",
        "design": "Cross-sectional Study",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "-",
        "checks": [True, True, True, True, True, True],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "Completed",
        "remarks": "ตีพิมพ์ในระดับนานาชาติเรียบร้อย รูปเล่มวิจัยเสร็จสมบูรณ์ 100%"
    },
    {
        "name": "พญ. กุลณัฐ เจษฎาพันธ์",
        "subspecialty": "Glaucoma",
        "title": "การเปลี่ยนแปลงของมุมตาหลังการหยอดยากลุ่ม Rho-kinase Inhibitor ในต้อหินมุมเปิด",
        "design": "Prospective Interventional Study",
        "advisor": "อ.พญ. ชยณัฐ อุยเจริญ",
        "co_advisor": "รศ.นพ. สุรชัย เมตตา",
        "checks": [True, True, True, True, True, False],
        "status_formula": '=IF(H{row}=1, "Completed", IF(N{row}=FALSE, "Behind Schedule", "On Track"))',
        "flag": "On Track",
        "remarks": "สถิติเสร็จแล้ว อยู่ระหว่างการตรวจทานบทคัดย่อภาษาอังกฤษร่วมกับ Advisor"
    },
]

build_cohort_sheet("R1_2572", "แพทย์ประจำบ้านชั้นปีที่ 1 (R1)", "พ.ศ. 2569", "พ.ศ. 2572", r1_timeline, r1_residents)
build_cohort_sheet("R2_2571", "แพทย์ประจำบ้านชั้นปีที่ 2 (R2)", "พ.ศ. 2568", "พ.ศ. 2571", r2_timeline, r2_residents)
build_cohort_sheet("R3_2570", "แพทย์ประจำบ้านชั้นปีที่ 3 (R3)", "พ.ศ. 2567", "พ.ศ. 2570", r3_timeline, r3_residents)

# -------------------------------------------------------------
# TAB 5: Config & Timeline
# -------------------------------------------------------------
ws_cfg = wb.create_sheet(title="Config_Timeline")
ws_cfg.views.sheetView[0].showGridLines = True

ws_cfg["A1"] = "ตารางกำหนดการแม่บทและค่าตัวเลือก (Master Timeline & System Configuration)"
ws_cfg["A1"].font = title_font
ws_cfg["A2"] = "ใช้สำหรับอ้างอิงช่วงเวลา Milestone, รายชื่อตัวเลือกประเภทงานวิจัย, และเกณฑ์การประเมินสถานะ"
ws_cfg["A2"].font = subtitle_font

# Table 1: Master Timeline
ws_cfg["A4"] = "1. กำหนดการ Milestone 6 ขั้นตอนตามหลักสูตรการฝึกอบรม (อิงตามภาพอ้างอิงของภาควิชา)"
ws_cfg["A4"].font = sub_header_font

cfg_headers = ["Milestone", "ชื่อขั้นตอน (Activity)", "ระยะเวลาเป้าหมาย R1 (เข้า 69)", "ระยะเวลาเป้าหมาย R2 (เข้า 68)", "ระยะเวลาเป้าหมาย R3 (เข้า 67)", "เกณฑ์ชี้วัดความสำเร็จ (Deliverables)"]
for c_idx, h in enumerate(cfg_headers, start=1):
    cell = ws_cfg.cell(row=5, column=c_idx, value=h)
    cell.font = header_font
    cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_header

timeline_matrix = [
    ("Milestone 1", "ทบทวนวรรณกรรมและกำหนดกรอบแนวคิด", "ก.ค. 2569 - ต.ค. 2569", "ก.ค. 2568 - ต.ค. 2568", "ก.ค. 2567 - ต.ค. 2567", "Literature Review Report + PICO Blueprint"),
    ("Milestone 2", "ยื่นโครงร่างวิจัยพิจารณาจริยธรรม (IRB)", "พ.ย. 2569 - ก.พ. 2570", "พ.ย. 2568 - ก.พ. 2569", "พ.ย. 2567 - ก.พ. 2568", "ใบรับรองการยื่นจริยธรรมวิจัย (IRB Submission Receipt)"),
    ("Milestone 3", "ขอทุนสนับสนุนงานวิจัยและยื่นแก้ไขพิจารณา", "มี.ค. 2570 - มิ.ย. 2570", "มี.ค. 2569 - มิ.ย. 2569", "มี.ค. 2568 - มิ.ย. 2568", "IRB Approval Letter + สัญญารับทุนวิจัย"),
    ("Milestone 4", "เก็บรวบรวมข้อมูลงานวิจัย (Data Collection)", "ก.ค. 2570 - ต.ค. 2570", "ก.ค. 2569 - ต.ค. 2569", "ก.ค. 2568 - ต.ค. 2568", "ฐานข้อมูลงานวิจัยที่ Cleaned แล้ว (CRF Database)"),
    ("Milestone 5", "วิเคราะห์ข้อมูลทางสถิติ (Statistical Analysis)", "พ.ย. 2570 - ม.ค. 2571", "พ.ย. 2569 - ม.ค. 2570", "พ.ย. 2568 - ม.ค. 2569", "ผลวิเคราะห์สถิติ Table 1, 2, 3 + p-values"),
    ("Milestone 6", "จัดทำรูปเล่มและเผยแพร่ผลงานทางวิชาการ", "ก.พ. 2571 - มิ.ย. 2572", "ก.พ. 2570 - มิ.ย. 2571", "ก.พ. 2569 - มิ.ย. 2570", "รูปเล่มวิจัยฉบับสมบูรณ์ / ตอบรับตีพิมพ์ (Publication)"),
]

for r_idx, row in enumerate(timeline_matrix, start=6):
    for c_idx, val in enumerate(row, start=1):
        cell = ws_cfg.cell(row=r_idx, column=c_idx, value=val)
        cell.font = data_font
        cell.border = border_thin
        if c_idx in [1, 3, 4, 5]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

# Table 2: Dropdown options
ws_cfg["A14"] = "2. รายการตัวเลือกสำหรับ Data Validation Dropdown"
ws_cfg["A14"].font = sub_header_font

ws_cfg["A15"] = "ประเภทงานวิจัย (Study Design)"
ws_cfg["A15"].font = header_font
ws_cfg["A15"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")

for idx, d in enumerate(designs, start=16):
    ws_cfg.cell(row=idx, column=1, value=d).font = data_font
    ws_cfg.cell(row=idx, column=1).border = border_thin

ws_cfg["C15"] = "สาขาย่อย (Subspecialties)"
ws_cfg["C15"].font = header_font
ws_cfg["C15"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")

subs = ["Glaucoma", "Retina", "Cornea & Refractive", "Pediatric Ophthalmology & Strabismus", "Neuro-Ophthalmology", "Oculoplastic & Reconstructive", "Uveitis", "Cataract & Comprehensive"]
for idx, s in enumerate(subs, start=16):
    ws_cfg.cell(row=idx, column=3, value=s).font = data_font
    ws_cfg.cell(row=idx, column=3).border = border_thin

ws_cfg["E15"] = "อาจารย์ที่ปรึกษา (Advisors)"
ws_cfg["E15"].font = header_font
ws_cfg["E15"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")

advisors = [
    "อ.พญ. ชยณัฐ อุยเจริญ",
    "รศ.นพ. สุรชัย เมตตา",
    "ผศ.นพ. วีระศักดิ์ แสงธรรม",
    "รศ.พญ. ศิริพร ทิพยรัตน์",
    "ผศ.พญ. ปัทมา วิเศษโภคัย",
    "รศ.นพ. ประพันธ์ ฤทธิ์เดช",
    "ผศ.พญ. สุดารัตน์ มีสุข",
    "อ.นพ. ธนกฤต ชัยชนะ"
]
for idx, adv in enumerate(advisors, start=16):
    ws_cfg.cell(row=idx, column=5, value=adv).font = data_font
    ws_cfg.cell(row=idx, column=5).border = border_thin

# Column dimensions for Config
ws_cfg.column_dimensions["A"].width = 28
ws_cfg.column_dimensions["B"].width = 38
ws_cfg.column_dimensions["C"].width = 28
ws_cfg.column_dimensions["D"].width = 28
ws_cfg.column_dimensions["E"].width = 30
ws_cfg.column_dimensions["F"].width = 44

# Adjust Dashboard column widths
dash_col_widths = {
    "A": 26, "B": 24, "C": 14, "D": 22, "E": 16, "F": 16, "G": 20, "H": 36, "I": 16, "J": 16, "K": 16, "L": 16
}
for col, width in dash_col_widths.items():
    ws_dash.column_dimensions[col].width = width

output_path = "/Users/CO/AI sandbox/projects/Research/Resident_Research_Progress_Tracker.xlsx"
wb.save(output_path)
print(f"Successfully generated: {output_path}")
