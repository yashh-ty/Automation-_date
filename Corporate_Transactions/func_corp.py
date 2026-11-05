import openpyxl
from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
from openpyxl.utils import get_column_letter
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

def first_report_header(ws,start_row_id:int):
    end_row_id=start_row_id+1

    ws.merge_cells(f"A{start_row_id}:A{end_row_id}")
    ws[f'A{start_row_id}'] = 'Sr No'

    ws.merge_cells(f'B{start_row_id}:B{end_row_id}')
    ws[f'B{start_row_id}'] = 'Contribution Details'

    ws.merge_cells(f'C{start_row_id}:C{end_row_id}')
    ws[f'C{start_row_id}'] = 'Type/Channel'


    ws.merge_cells(f'D{start_row_id}:F{start_row_id}')
    ws[f'D{start_row_id}'] = 'Count'

    ws.merge_cells(f'G{start_row_id}:H{start_row_id}')
    ws[f'G{start_row_id}'] = 'Amount'

    ws.merge_cells(f'I{start_row_id}:K{start_row_id}')
    ws[f'I{start_row_id}'] = 'Count'

    ws.merge_cells(f'L{start_row_id}:M{start_row_id}')
    ws[f'L{start_row_id}'] = 'Amount'

    subheaders = {f"D{end_row_id}": 'Successful', f"E{end_row_id}": "Failed", f"F{end_row_id}": "%Failure",
                  f"G{end_row_id}": 'Successful', f"H{end_row_id}": "Failed",
                  f"I{end_row_id}": 'Successful', f"J{end_row_id}": "Failed", f"K{end_row_id}": "%Failure",
                  f"L{end_row_id}": "Successful", f"M{end_row_id}": "Failed"

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
