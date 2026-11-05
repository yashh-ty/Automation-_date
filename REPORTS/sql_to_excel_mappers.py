

def excel_mappers(report_data,mapper_excel):
    final_output={}
    for row_idx, excel_tuple in report_data.items():

        # Step A: Find the matching SQL tuple using Level 2 Mapper
        sql_tuple = mapper_excel.get(excel_tuple)

        final_output[row_idx]=sql_tuple
    return final_output





