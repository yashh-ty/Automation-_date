import datetime
import time
from MB.mb_headers import *
from openpyxl.cell.cell import MergedCell
from REPORTS.utlis import *
from openpyxl.utils import get_column_letter
start_time = time.time()

exception_rows=[5,24,27,44,46,51]



def mb_generate_report(ws):
    headers_map = {}
    current_row = 5
    # 1. Main Headers
    date = datetime.date.today().strftime(format="%d-%m-%Y")
    left_header = date
    right_header = f"MTD (1st to {date.split('-')[0]} July)"
    top_header(ws, "Summary Report", left_header, right_header)

    first_report_header(ws, 3)

    # banner=[(5,'A',"N","PAY"),(24,"A","N",None),(27,"A","N","INVEST"),(46,"A","N",None),(51,"A","N",None),(61,"A","D","Total")]

    dict={5:{"start_cell":'A5',"end_cell":'N5',"value":"PAY"},
         24:{"start_cell":'A24',"end_cell":'N24',"value":None},
         27:{"start_cell":'A27',"end_cell":'N27',"value":"INVEST"},
         46:{"start_cell" :'A46',"end_cell":'N46',"value":None},
         51:{"start_cell":'A51',"end_cell":'N51',"value":None}}


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




    second_report_header(ws,current_row)
    current_row+=2

    for j in section2:
        while current_row in exception_rows:
            sep_blue_line(ws, dict[current_row]["start_cell"], dict[current_row]["end_cell"],dict[current_row]["value"])
            current_row += 1
        b_val, c_val, d_val=j[0],j[1],j[2]
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

        current_row+=1

    # total_count_report
    target_row = current_row
    sep_blue_line(ws,f'A{target_row}',f'D{target_row}','Total')

    # Current-->Count--->Sucessfull
    ws[f"E{target_row}"] = f"=SUM(E6:E23, E25:E26, E28:E43, E46:E50, E52:E60)"

    # Current-->Count--->Failed
    ws[f"F{target_row}"] = f"=SUM(F6:F23, F25:F26, F28:F43, F46:F50, F52:F60)"

    # Current-->Amount--->Sucessfull
    ws[f"H{target_row}"] = f"=SUM(H6:H23, H25:H26, H28:H43, H46:H50, H52:H60)"

    # Current-->Amount--->Failed
    ws[f"I{target_row}"] = f"=SUM(I6:I23, I25:I26, I28:I43, I46:I50, I52:I60)"

    # MTD-->Count--->Successfull
    ws[f"J{target_row}"] = f"=SUM(J6:J23, J25:J26, J28:J43, J46:J50, J52:J60)"

    # MTD-->Count--->Failed
    ws[f"K{target_row}"] = f"=SUM(K6:K23, K25:K26, K28:K43, K46:K50, K52:K60)"

    # MTD-->Amount--->Successful
    ws[f"M{target_row}"] = f"=SUM(M6:M23, M25:M26, M28:M43, M46:M50, M52:M60)"

    # MTD-->Amount--->Failed
    ws[f"N{target_row}"] = f"=SUM(N6:N23, N25:N26, N28:N43, N46:N50, N52:N60)"

    #Total Failure %
    ws[f"G{target_row}"] = (f"=IF((E{target_row}+F{target_row})=0, 0, (F{target_row}/(E{target_row}+F{target_row}))*100)")
    ws[f"G{target_row}"].number_format = "0.00"

    ws[f"L{target_row}"] = (f"=IF((K{target_row}+J{target_row})=0, 0, (K{target_row}/(K{target_row}+J{target_row}))*100)")
    ws[f"L{target_row}"].number_format = "0.00"





    # Border and  center alignment
    for rows in range(2, target_row+1):
        for columns in range(1, 14 + 1):
            cell = ws.cell(row=rows, column=columns)
            # if isinstance(cell, MergedCell):
            #     continue
            cell.alignment = center_align
            cell.border = thin_border


    #Column Spacing Adjustment
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
    index=1
    for row_idx in range(5, ws.max_row ):
        cell = ws.cell(row=row_idx, column=1)
        if isinstance(cell, MergedCell) or row_idx in exception_rows:
            continue
        cell.value = index
        index+=1



    vertical_merge(ws,"B6","B12")
    vertical_merge(ws,"B13",'B16')
    vertical_merge(ws,"B19",'B22')
    vertical_merge(ws,"B30",'B32')


    return headers_map





# file_path=r"C:\Users\Alchemy\Desktop\reports\TRANSACTION3_REPORT.xlsx"
# wb = openpyxl.Workbook()
# ws_mb=wb.active
# ws_mb.title = "MB-Transactions"
#
# print(mb_generate_report(ws_mb))
# wb.save(file_path)



