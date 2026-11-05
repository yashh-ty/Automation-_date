
from REPORTS.utlis import *
from .func_corp import *
from .corp_headers  import *
from openpyxl.cell.cell import MergedCell
import datetime
def corp_generate_report(ws):
    headers_map = {}
    current_row = 6
    exception_rows=[3,4,15]
    # 1. Main Headers
    date = datetime.date.today().strftime(format="%d-%m-%Y")
    left_header = date
    right_header = f"MTD (1st to {date.split('-')[0]} July)"
    top_header(ws, "Summary Report", left_header, right_header)

    first_report_header(ws,4)

    dict={
        5: {"start_cell": 'A5', "end_cell": 'N5', "value": "PAY"}
    }

    for i in section:
        while current_row in exception_rows:
            sep_blue_line(ws,dict[current_row]["start_cell"],dict[current_row]["end_cell"],dict[current_row]["value"])
            current_row+=1
        b_val, c_val= i[0], i[1]
        ws[f"B{current_row}"] = b_val
        ws[f"C{current_row}"] = c_val

        # Apply Row Percentage Formulas
        ws[f"F{current_row}"] = (
            f"=IF((D{current_row}+E{current_row})=0, 0, (E{current_row}/(D{current_row}+E{current_row}))*100)"
        )
        ws[f"F{current_row}"].number_format = "0.00"

        ws[f"K{current_row}"] = (
            f"=IF((I{current_row}+J{current_row})=0, 0, (J{current_row}/(I{current_row}+J{current_row}))*100)"
        )
        ws[f"K{current_row}"].number_format = "0.00"

        headers_map[current_row] = (b_val, c_val)

        current_row += 1


    target_row=current_row
    sep_blue_line(ws, f'A{target_row}', f'C{target_row}', 'Total')

    # Current-->Count--->Sucessfull
    ws[f"D{target_row}"] = f"=SUM(D6:D14)"

    # Current-->Count--->Failed
    ws[f"E{target_row}"] = f"=SUM(E6:E14)"

    # Current-->Amount--->Sucessfull
    ws[f"G{target_row}"] = f"=SUM(G6:G14)"

    # Current-->Amount--->Failed
    ws[f"H{target_row}"] = f"=SUM(H6,H14)"

    # MTD-->Count--->Successfull
    ws[f"I{target_row}"] = f"=SUM(I6:I14)"

    # MTD-->Count--->Failed
    ws[f"J{target_row}"] = f"=SUM(J6:J14)"

    # MTD-->Amount--->Successful
    ws[f"L{target_row}"] = f"=SUM(L6:L14)"

    # MTD-->Amount--->Failed
    ws[f"M{target_row}"] = f"=SUM(M6:M14)"

    # Total Failure %
    ws[f"F{target_row}"] = (f"=IF((D{target_row}+E{target_row})=0, 0, (E{target_row}/(D{target_row}+E{target_row}))*100)")
    ws[f"F{target_row}"].number_format = "0.00"

    ws[f"K{target_row}"] = (f"=IF((I{target_row}+J{target_row})=0, 0, (J{target_row}/(I{target_row}+J{target_row}))*100)")
    ws[f"K{target_row}"].number_format = "0.00"


    vertical_merge(ws,"B6","B11")
    vertical_merge(ws,"B12","B14")

    # Border and  center alignment
    for rows in range(2, ws.max_row + 1):
        for columns in range(1, 14 + 1):
            cell = ws.cell(row=rows, column=columns)
            cell.alignment = center_align
            cell.border = thin_border

    # Column Spacing Adjustment
    for col in ws.columns:
            max_len = 0
            for cell in col:

                # Skip title rows (1-4) and formula strings (=...)
                val = str(cell.value or "")
                if cell.row < 5 or val.startswith("="):
                    continue
                max_len = max(max_len, len(val))

            col_letter = get_column_letter(col[0].column)
            # Set width between minimum 12 and max data length + padding
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Starting from row 1 down to the last row with data
    index = 1
    for row_idx in range(5, ws.max_row):
        cell = ws.cell(row=row_idx, column=1)
        if isinstance(cell, MergedCell) or row_idx in exception_rows:
            continue
        cell.value = index
        index += 1

    return headers_map





