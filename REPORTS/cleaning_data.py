import csv

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



def matching(report_mapper_excel,daily_data,monthly_data):
    final_output={}
    for i in report_mapper_excel.values():
            if i==None:
                final_output[i] =('0','0','0','0','0','0','0','0')
            elif i in daily_data and i not in monthly_data:
                final_output[i] = daily_data[i] +('0','0','0','0')
            elif i not in daily_data and i  in monthly_data:
                final_output[i] = ('0','0','0','0')+monthly_data[i]
            else:
                final_output[i] = daily_data[i] + monthly_data[i]


    return final_output


##Remember that It is only for 'MB-Transactions' or 'IB-Transactions' Sheet
def write_mapped_dict_to_excel(ws,row_to_sql_map, sql_data):


    # Explicit column targets: E (5), F (6), H (8), I (9)
    col_mapping = [5, 6, 8,9,10,11,13,14]

    for row_idx, sql_tuple in row_to_sql_map.items():
        # Get SQL metrics tuple (default to 0s if missing)
        metrics = sql_data.get(sql_tuple, ('0', '0', '0', '0','0','0','0','0'))

        # Write each metric to its corresponding target column
        for i, val in enumerate(metrics):
            target_col = col_mapping[i]  # Picks column 5, 6, 8, or 9
            clean_val = float(val) if '.' in str(val) else int(val)
            ws.cell(row=row_idx, column=target_col, value=clean_val)






def corp_write_mapped_dict_to_excel(ws,row_to_sql_map, sql_data):



    # Explicit column targets: E (5), F (6), H (8), I (9)
    col_mapping = [4, 5, 7,8,9,10,12,13]

    for row_idx, sql_tuple in row_to_sql_map.items():
        # Get SQL metrics tuple (default to 0s if missing)
        metrics = sql_data.get(sql_tuple, ('0', '0', '0', '0','0','0','0','0'))

        # Write each metric to its corresponding target column
        for i, val in enumerate(metrics):
            target_col = col_mapping[i]  # Picks column 5, 6, 8, or 9
            clean_val = float(val) if '.' in str(val) else int(val)
            ws.cell(row=row_idx, column=target_col, value=clean_val)


