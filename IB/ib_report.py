from openpyxl.utils import get_column_letter
from REPORTS.utlis import *
import datetime
from openpyxl.cell.cell import MergedCell
from IB.ib_headers import section1,section2
exception_rows=[5,28,31,44,45,46,51,62]
def ib_generate_report(ws):
    headers_map={}
    current_row=5
    # 1. Main Headers
    date = datetime.date.today().strftime(format="%d-%m-%Y")
    left_header = date
    right_header = f"MTD (1st to {date.split('-')[0]} July)"
    top_header(ws, "Summary Report", left_header, right_header)

    first_report_header(ws,3)

    dict = { 5: {"start_cell": 'A5', "end_cell": 'N5', "value": "PAY"},
            28: {"start_cell": 'A28', "end_cell": 'N28', "value": None},
            31: {"start_cell": 'A31', "end_cell": 'N31', "value": "INVEST"},
            46: {"start_cell": 'A46', "end_cell": 'N46', "value": "SAVE"},
            51: {"start_cell": 'A51', "end_cell": 'N51', "value": None}
            }

    for i in section1:
        while current_row in exception_rows:
            sep_blue_line(ws,dict[current_row]["start_cell"],dict[current_row]["end_cell"],dict[current_row]["value"])
            current_row+=1

        b_val, c_val, d_val=i[0],i[1],i[2]
        ws[f"B{current_row}"] = b_val
        ws[f"C{current_row}"] = c_val
        ws[f"D{current_row}"] = d_val

        # Apply Row Percentage Formulas
        ws[f"G{current_row}"] = (
        f"=IF((E{current_row}+F{current_row})=0, 0, (F{current_row}/(E{current_row}+F{current_row}))*100)"
        )
        ws[f"G{current_row}"].number_format = "0.00"

        ws[f"L{current_row}"] = (
        f"=IF((J{current_row}+K{current_row})=0, 0, (K{current_row}/(J{current_row}+K{current_row}))*100)"
        )
        ws[f"L{current_row}"].number_format = "0.00"

        headers_map[current_row]=(b_val,c_val,d_val)

        current_row+=1

    second_report_header(ws, current_row)
    current_row += 2

    for j in section2:
        while current_row in exception_rows:
            sep_blue_line(ws, dict[current_row]["start_cell"], dict[current_row]["end_cell"],
                          dict[current_row]["value"])
            current_row += 1
        b_val, c_val, d_val = j[0], j[1], j[2]
        ws[f"B{current_row}"] = b_val
        ws[f"C{current_row}"] = c_val
        ws[f"D{current_row}"] = d_val
        # Apply Row Percentage Formulas
        ws[f"G{current_row}"] = (
            f"=IF((E{current_row}+F{current_row})=0, 0, (F{current_row}/(E{current_row}+F{current_row}))*100)"
        )
        ws[f"G{current_row}"].number_format = "0.00"

        ws[f"L{current_row}"] = (
            f"=IF((J{current_row}+K{current_row})=0, 0, (K{current_row}/(J{current_row}+K{current_row}))*100)"
        )
        ws[f"L{current_row}"].number_format = "0.00"

        headers_map[current_row] = (b_val, c_val, d_val)

        current_row += 1

    # total_count_report
    target_row = current_row
    sep_blue_line(ws, f'A{target_row}', f'D{target_row}', 'Total')

    # Current-->Count--->Sucessfull
    ws[f"E{target_row}"] = f"=SUM(E6:E27, E29:E30, E32:E43, E47:E50, E52:E61)"

    # Current-->Count--->Failed
    ws[f"F{target_row}"] = f"=SUM(F6:F27, F29:F30, F32:F43, F47:F50, F52:F61)"

    # Current-->Amount--->Sucessfull
    ws[f"H{target_row}"] = f"=SUM(H6:H27, H29:H30, H32:H43, H47:H50, H52:H61)"

    # Current-->Amount--->Failed
    ws[f"I{target_row}"] = f"=SUM(I6:I27, I29:I30, I32:I43, I47:I50, I52:I61)"

    # MTD-->Count--->Successfull
    ws[f"J{target_row}"] = f"=SUM(J6:J27, J29:J30, J32:J43, J47:J50, J52:J61)"

    # MTD-->Count--->Failed
    ws[f"K{target_row}"] = f"=SUM(K6:K27, K29:K30, K32:K43, K47:K50, K52:K61)"

    # MTD-->Amount--->Successful
    ws[f"M{target_row}"] = f"=SUM(M6:M27, M29:M30, M32:M43, M47:M50, M52:M61)"

    # MTD-->Amount--->Failed
    ws[f"N{target_row}"] = f"=SUM(N6:N27, N29:N30, N32:N43, N47:N50, N52:N61)"

    # Total Failure %
    ws[f"G{target_row}"] = (f"=IF((E{target_row}+F{target_row})=0, 0, (F{target_row}/(E{target_row}+F{target_row}))*100)")
    ws[f"G{target_row}"].number_format = "0.00"

    ws[f"L{target_row}"] = (f"=IF((K{target_row}+J{target_row})=0, 0, (K{target_row}/(K{target_row}+J{target_row}))*100)")
    ws[f"L{target_row}"].number_format = "0.00"

    vertical_merge(ws, "B6", 'B12')
    vertical_merge(ws, "B14", "B17")
    vertical_merge(ws, "B22", 'B27')
    vertical_merge(ws, "B34", 'B36')
    vertical_merge(ws, "B56", 'B61')

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



# file_path=r"C:\Users\Alchemy\Desktop\reports\TRANSACTION3_REPORT.xlsx"
# wb = openpyxl.Workbook()
# ws_ib=wb.active
# ws_ib.title = "IB-Transactions"
#
# print(ib_generate_report(ws_ib))
# wb.save(file_path)



