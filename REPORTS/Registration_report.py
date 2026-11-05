


import csv
import openpyxl
from openpyxl.styles import Font
from REPORTS.utlis import (
    center_align,
    thin_border,
    yellow_fill,
)

def reg_report(ws, Count_data, Corp_IB_Count, Status_count):

    # Load CSV data
    with open(fr"{Count_data}") as csvfile:
        count_data = list(csv.reader(csvfile, delimiter=","))[1:]

    with open(fr"{Corp_IB_Count}") as csvfile:
        corp_ib_count = list(csv.reader(csvfile, delimiter=","))[1:]

    with open(fr"{Status_count}") as csvfile:
        registration_status = list(csv.reader(csvfile, delimiter=","))[1:]



    bold_font = Font(bold=True)

    def apply_header(row, col, value):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = yellow_fill
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border
        return cell

    # ------------------ SECTION 1: Registration Details ------------------
    report1_start = 1
    apply_header(report1_start, 1, 'Registration Details')
    apply_header(report1_start, 2, 'Total')
    apply_header(report1_start, 3, 'Today')

    for i, j in enumerate(count_data, start=report1_start + 1):
        for col_idx, val in enumerate(j[:3], start=1):
            cell = ws.cell(row=i, column=col_idx, value=val if col_idx == 1 else int(val))
            cell.border = thin_border
            # cell.alignment = center_align

    # ------------------ SECTION 2: Corp_IB_COUNT ------------------
    report2_start = report1_start + len(count_data) + 3 # dynamic spacing

    apply_header(report2_start, 1, 'Corp_IB Count')
    apply_header(report2_start, 2, 'Total')
    apply_header(report2_start, 3, 'Today')

    for i, j in enumerate(corp_ib_count, start=report2_start + 1):
        for col_idx, val in enumerate(j[:3], start=1):
            cell = ws.cell(row=i, column=col_idx, value=val if col_idx == 1 else int(val))
            cell.border = thin_border
            # cell.alignment = center_align

    # Add Total Row
    total_row = report2_start + len(corp_ib_count) + 1
    cell_lbl = ws.cell(row=total_row, column=1, value='Total')
    cell_lbl.font = bold_font
    cell_lbl.border = thin_border

    cell_b = ws.cell(row=total_row, column=2, value=f"=SUM(B{report2_start+1}:B{total_row-1})")
    cell_b.border = thin_border


    cell_c = ws.cell(row=total_row, column=3, value=f"=SUM(C{report2_start+1}:C{total_row-1})")
    cell_c.border = thin_border


    # ------------------ SECTION 3: REGISTRATION STATUS ------------------
    report3_start = total_row + 3 # dynamic spacing

    apply_header(report3_start, 1, 'Status')
    apply_header(report3_start, 2, 'Count')

    for i, j in enumerate(registration_status, start=report3_start + 1):
        c1 = ws.cell(row=i, column=1, value=j[0])
        c2 = ws.cell(row=i, column=2, value=int(j[1]))
        c1.border, c2.border = thin_border, thin_border

    # Auto-adjust column widths based on content length
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter  # Get column name ('A', 'B', 'C')
        for cell in col:
            if cell.value:
                # Get length of string representation of cell value
                val_str = str(cell.value)
                max_len = max(max_len, len(val_str))

        # Set column width with a little extra padding
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


    # wb.save(excel_file_path)

# reg_report(r"C:\Users\Alchemy\Desktop\21-09-2026 Transaction Report.xlsx",r"C:\Users\Alchemy\Desktop\21-09-2026 Report\COUNT_DATA",r"C:\Users\Alchemy\Desktop\21-09-2026 Report\CORP_IB_COUNT",r"C:\Users\Alchemy\Desktop\21-09-2026 Report\REGISTRATION_STATUS")