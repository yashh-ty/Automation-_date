
import csv

from REPORTS.utlis import (
    body_font,
    center_align,
    section_fill,
    thin_border,
    yellow_fill,
)


def state_tax(target_row,ws, data_path):
    with open(fr"{data_path}") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        data = list(reader)[1:]


    row1 = target_row
    row2 = target_row + 1
    row3 = target_row + 2

    # 1. Merge regions FIRST
    ws.merge_cells(
        start_row=row1, start_column=1, end_row=row1, end_column=5
    )  # (State Tax)
    ws.merge_cells(
        start_row=row2, start_column=1, end_row=row3, end_column=1
    )  # B2:B3 (State)
    ws.merge_cells(
        start_row=row2, start_column=2, end_row=row2, end_column=3
    )  # C2:D2 (Daily)

    ws.merge_cells(start_row=row2, start_column=4, end_row=row2, end_column=5)
    # Monthly

    # 2. Write values to primary top-left cells
    ws.cell(row=row1, column=1, value="State Tax")
    ws.cell(row=row2, column=1, value="State")

    ws.cell(row=row2, column=2, value="Daily")
    ws.cell(row=row3, column=2, value="Count")
    ws.cell(row=row3, column=3, value="Amount")

    ws.cell(row=row2, column=4, value="Month")
    ws.cell(row=row3, column=4, value="Count")
    ws.cell(row=row3, column=5, value="Amount")

    # 3. Format Row 1 (Light Blue Block: B to F)
    for col in range(1, 5 + 1):
        cell = ws.cell(row=row1, column=col)
        cell.fill = section_fill
        cell.font = body_font
        cell.alignment = center_align
        cell.border = thin_border

    # 4. Format Rows 2 & 3 (Yellow Block: B to F across BOTH rows)
    for r in range(row2, row3 + 1):
        for c in range(1, 5 + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = yellow_fill
            cell.font = body_font
            cell.alignment = center_align
            cell.border = thin_border

    # Data filling
    start_row = row3 + 1
    for i in data:
        ws.cell(row=start_row, column=1, value=i[0])
        ws.cell(row=start_row,column=2,value=int(i[1]) if i[1].isdigit() else i[1])
        ws.cell(row=start_row,column=3,value=float(i[2]) if "." in i[2] else int(i[2]))
        ws.cell(row=start_row,column=4,value=int(i[3]) if i[3].isdigit() else i[3])
        ws.cell(row=start_row,column=5,value=float(i[4]) if "." in i[4] else int(i[4]))
        start_row += 1

    # Formatting data rows starting AFTER row3 (prevents MergedCell error)
    for i in range(row3 + 1, start_row):
        for j in range(1, 6):
            cell = ws.cell(row=i, column=j)
            cell.border = thin_border
            cell.alignment = center_align




# state_tax(64,'IB-Transactions',r"C:\Users\Alchemy\Desktop\21-09-2026 Transaction Report.xlsx",r"C:\Users\Alchemy\Desktop\21-09-2026 Report\IB-STATE TAX")
#
#
# state_tax(17,'Corporate-Transactions',r"C:\Users\Alchemy\Desktop\21-09-2026 Transaction Report.xlsx",r"C:\Users\Alchemy\Desktop\21-09-2026 Report\CORP-STATE TAX")



