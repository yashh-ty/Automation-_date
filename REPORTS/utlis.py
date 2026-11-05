from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
from openpyxl.utils.cell import coordinate_to_tuple


# Define Styles
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")


header_font=Font(name='Calibri',size=11)
body_font=Font(name='Calibri',size=11,bold=True)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center")
right_align = Alignment(horizontal="right", vertical="center")

thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
)


def top_header(ws,main_header,left_header,right_header):
    ws.merge_cells("E1:I1")
    cell = ws['E1']
    cell.value = f"{main_header}"
    cell.font = Font(bold=True)
    cell.alignment = center_align

    ws.merge_cells("E2:I2")
    cell = ws['E2']
    cell.value = f"{left_header}"
    cell.font = Font(bold=True)
    cell.alignment = center_align
    cell.border = thin_border
    for col in range(5, 10):
        ws.cell(row=2, column=col).border = thin_border

    ws.merge_cells("J2:N2")
    cell = ws["J2"]
    cell.value = f"{right_header}"
    cell.font = Font(bold=True)

    cell.alignment = center_align
    cell.border = thin_border
    for col in range(10, 15):
        ws.cell(row=2, column=col).border = thin_border


# give start_cell value ---> 'A5'
# end_cell_value -->'B5'

def sep_blue_line(ws,start_cell:str,end_cell:str,value=None):
    ws.merge_cells(f"{start_cell}:{end_cell}")
    target_cell=ws[f"{start_cell}"]
    if value!=None:
        target_cell.value=value
    target_cell.alignment = left_align
    target_cell.font = header_font

    start_row, _ = coordinate_to_tuple(start_cell)

    for col in range(1, 15):
        cell = ws.cell(row=start_row, column=col)
        cell.fill = section_fill
        cell.border = thin_border
        cell.alignment = center_align


#PASS THE ROW_ID FROM WHICH YOU WANT TO GIVE THE FIRST HEADER
def first_report_header(ws,start_row_id:int):
    end_row_id=start_row_id+1

    ws.merge_cells(f"A{start_row_id}:A{end_row_id}")
    ws[f'A{start_row_id}'] = 'Sr No'

    ws.merge_cells(f'B{start_row_id}:B{end_row_id}')
    ws[f'B{start_row_id}'] = 'Contribution Details'

    ws.merge_cells(f'C{start_row_id}:C{end_row_id}')
    ws[f'C{start_row_id}'] = 'Type/Channel'

    ws.merge_cells(f'D{start_row_id}:D{end_row_id}')
    ws[f'D{start_row_id}'] = 'SubType'

    ws.merge_cells(f'E{start_row_id}:G{start_row_id}')
    ws[f'E{start_row_id}'] = 'Count'

    ws.merge_cells(f'H{start_row_id}:I{start_row_id}')
    ws[f'H{start_row_id}'] = 'Amount'

    ws.merge_cells(f'J{start_row_id}:L{start_row_id}')
    ws[f'J{start_row_id}'] = 'Count'

    ws.merge_cells(f'M{start_row_id}:N{start_row_id}')
    ws[f'M{start_row_id}'] = 'Amount'

    subheaders = {f"E{end_row_id}": 'Successful', f"F{end_row_id}": "Failed", f"G{end_row_id}": "%Failure",
                  f"H{end_row_id}": 'Successful', f"I{end_row_id}": "Failed",
                  f"J{end_row_id}": 'Successful', f"K{end_row_id}": "Failed", f"L{end_row_id}": "%Failure",
                  f"M{end_row_id}": "Successful", f"N{end_row_id}": "Failed"

                  }

    for row, header in subheaders.items():
        ws[row] = header

    for row in range(start_row_id, end_row_id+1):
        for column in range(1, 15):
            cell = ws.cell(row=row, column=column)
            cell.fill = yellow_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border



def second_report_header(ws,start_row_id:int):
    end_row_id=start_row_id+1
    
    ws.merge_cells(f"A{start_row_id}:A{end_row_id}")

    ws.merge_cells(f"B{start_row_id}:B{end_row_id}")
    ws[f"B{start_row_id}"] = 'Account opened/Accounts closed/Policies created'

    ws.merge_cells(f"C{start_row_id}:C{end_row_id}")
    ws[f"C{start_row_id}"] = "Type"

    ws.merge_cells(f"D{start_row_id}:D{end_row_id}")
    ws[f"D{start_row_id}"] = "Sub-Type"

    ws.merge_cells(f"E{start_row_id}:G{start_row_id}")
    ws[f"E{start_row_id}"] = "Count"

    ws.merge_cells(f"H{start_row_id}:I{start_row_id}")
    ws[f"H{start_row_id}"] = "Amount"

    ws.merge_cells(f"J{start_row_id}:L{start_row_id}")
    ws[f"J{start_row_id}"] = "Count"

    ws.merge_cells(f"M{start_row_id}:N{start_row_id}")
    ws[f"M{start_row_id}"] = "Amount"

    subheaders = {f"E{end_row_id}": 'Successful', f"F{end_row_id}": "Failed", f"G{end_row_id}": "%Failure",
                  f"H{end_row_id}": 'Successful', f"I{end_row_id}": "Failed",
                  f"J{end_row_id}": 'Successful', f"K{end_row_id}": "Failed", f"L{end_row_id}": "%Failure",
                  f"M{end_row_id}": "Successful", f"N{end_row_id}": "Failed"
                  }

    for row, header in subheaders.items():
        ws[row] = header

    for row in range(start_row_id, end_row_id + 1):
        for column in range(1, 14 + 1):
            cell = ws.cell(row=row, column=column)
            cell.fill = yellow_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

def vertical_merge(ws,start_cell,end_cell):
    ws.merge_cells(f"{start_cell}:{end_cell}")







    

