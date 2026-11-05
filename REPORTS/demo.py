from row_headers_excel import ib_data
from sql_to_excel_mappers import ib_mapper_excel
import csv
import openpyxl

def fix_null(val):
    return 0 if val is None or str(val).strip().upper() in ("NULL", "") else val

def formatting_data(location):
    with open(fr'{location}') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        reader=list(reader)[1::]
        clean_reader = [
            [row[0], row[1], fix_null(row[2]), fix_null(row[3]), fix_null(row[4]), fix_null(row[5])]
            for row in reader
        ]

        data={(i[0],i[1]):(i[2],i[4],i[3],i[5]) for i in clean_reader}
        return data


daily_data=formatting_data(r"C:\Users\Alchemy\Desktop\29 REPORT\29_all_day_trnx_ib\ib_day_yash.csv")
monthly_data=formatting_data(r"C:\Users\Alchemy\Desktop\29 REPORT\29_all_day_trnx_ib\ib_month_yash.csv")
# print(daily_data)

#
# for (i,j) in daily_data.items():
#     print((i,j),sep='\n')

def matching(daily_data,monthly_data):
    final_output={}
    for i in daily_data:
        if i in monthly_data:
            final_output[i]=daily_data[i]+monthly_data[i]
    return final_output

sql_data=matching(daily_data,monthly_data)

# print(len(data))--->44

# ib_mapper_excel={('Fund Transfer', 'IMPS', ''): ("Fund transfer","IMPS"),
#         ('Fund Transfer', 'NEFT', ''): ("Fund transfer","NEFT"),
#         ('Fund Transfer', 'RTGS', ''): ("Fund transfer","RTGS"),
#         ('Fund Transfer', 'Within Bank', ''): ("Fund transfer","WITHINBOI")}

# ib_data={6: ('Fund Transfer', 'IMPS', ''),
#  7: ('Fund Transfer', 'NEFT', ''),
#  8: ('Fund Transfer', 'RTGS', ''),
#  9: ('Fund Transfer', 'Within Bank', '')

def excel_mappers(sql_data):
    final_output={}
    for row_idx, excel_tuple in ib_data.items():
        # Step A: Find the matching SQL tuple using Level 2 Mapper
        sql_tuple = ib_mapper_excel.get(excel_tuple)
        final_output[row_idx]=sql_tuple

    return final_output


final_data=excel_mappers(sql_data)
print(len(sql_data))
print(len(final_data))
#
# for (i,j) in final_data.items():
#     print((i,j),sep='\n')


def write_mapped_dict_to_excel(file_path,row_to_sql_map, sql_data):

    wb = openpyxl.load_workbook(file_path)
    ws = wb["IB-Transactions"]  # Or specify worksheet name: wb["Report"]

    # Explicit column targets: E (5), F (6), H (8), I (9)
    col_mapping = [5, 6, 8,9,10,11,13,14]

    for row_idx, sql_tuple in row_to_sql_map.items():
        # Get SQL metrics tuple (default to 0s if missing)
        metrics = sql_data.get(sql_tuple, ('10', '10', '10', '10','10','10','10','10'))
        # if metrics==None:
        #     metrics=('10', '10', '10', '10','10','10','10')
        print(metrics)


        # Write each metric to its corresponding target column
        for i, val in enumerate(metrics):
            target_col = col_mapping[i]  # Picks column 5, 6, 8, or 9
            clean_val = float(val) if '.' in str(val) else int(val)
            ws.cell(row=row_idx, column=target_col, value=clean_val)

    wb.save(file_path)
    print(f"Excel report updated successfully! Saved to: {file_path}")

write_mapped_dict_to_excel(r"C:\Users\Alchemy\Desktop\TRANSACTION_REPORT.xlsx",final_data,sql_data)

# def matching(daily_data,monthly_data):
#     final_output={}
#     for i in daily_data:
#         if i in monthly_data:
#             final_output[i]=daily_data[i]+monthly_data[i]
#         else:
#             final_output[i]=('0', '0', '0', '0', '0', '0', '0', '0')
#     return final_output
#
#
#
# sql_data=matching(daily_data,monthly_data)
#
# print(excel_mappers())

# ib_mapper_excel={('Fund Transfer', 'IMPS', ''): ("Fund transfer","IMPS"),
#         ('Fund Transfer', 'NEFT', ''): ("Fund transfer","NEFT"),
#         ('Fund Transfer', 'RTGS', ''): ("Fund transfer","RTGS"),
#         ('Fund Transfer', 'Within Bank', ''): ("Fund transfer","WITHINBOI")}
#
# ib_data={6: ('Fund Transfer', 'IMPS', ''),
#  7: ('Fund Transfer', 'NEFT', ''),
#  8: ('Fund Transfer', 'RTGS', ''),
#  9: ('Fund Transfer', 'Within Bank', '')}

# demo_data={}
# for row,data in ib_data.items():
#     sql_tuple = ib_mapper_excel.get(data)
#     if sql_tuple==None:
#         demo_data[row]=('0', '0', '0', '0','0','0','0','0')
#     else:
#         demo_data[row]=sql_tuple
#
# for (i,j) in demo_data.items():
#     print((i,j),sep='\n')


import openpyxl

# wb = openpyxl.load_workbook(
#     r"C:\Users\Alchemy\Desktop\TRANSACTION_REPORT-04-09-2026.xlsx"
# )
# ws = wb["MB-Transactions"]
#
#
# # Print all merged cell ranges in this sheet
# print(ws.merged_cells.ranges)

