import os

import logging
import openpyxl

from IB import ib_report
from MB import mb_report
from Corporate_Transactions import corp_transactions
from .cleaning_data import *
from .sql_to_excel_mappers import excel_mappers
from .mapper import mb_mapper_excel, ib_mapper_excel, corp_mapper_excel
from .row_headers_excel import mb_data, corp_data, ib_data
from .state_tax_report import state_tax
from .Registration_report import reg_report
from .report_date import prompt_report_date

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("ReportTransformer")

REPORTS_DIR = os.path.abspath(r"/home/yashraj/reports")


def report_generation():
    while True:
        print("==================================================")
        print("              REPORT GENERATION                   ")
        print("==================================================")
        print("\nNOTE: Ensure 01_extractor.py has finished dumping raw CSVs")
        print("into the staging folder before proceeding.\n")

        report_date = prompt_report_date()
        csv_directory = os.path.abspath(r"/home/yashraj/reports/night_staging_csv")
        report_name = "Transaction_Report.xlsx"

        os.makedirs(REPORTS_DIR, exist_ok=True)
        final_excel_path = os.path.join(REPORTS_DIR, report_name)

        logger.info(f"Processing staging data from: {csv_directory}")

        # Constructing explicit input file paths safely
        mb_daily_path = os.path.join(csv_directory, "MB QUERY DAILY.csv")
        mb_monthly_path = os.path.join(csv_directory, "MB MONTHLY QUERY.csv")

        ib_daily_path = os.path.join(csv_directory,  "IB QUERY DAILY.csv")
        ib_monthly_path = os.path.join(csv_directory, "IB MONTHLY QUERY.csv")

        corp_daily_path = os.path.join(csv_directory, "CORP QUERY DAILY.csv")
        corp_monthly_path = os.path.join(csv_directory, "CORP QUERY MONTHLY.csv")

        count_data = os.path.join(csv_directory, "REGISTRATION QUERY 1.csv")
        corp_ib_count = os.path.join(csv_directory,  "REGISTRATION QUERY 2.csv")
        status_count = os.path.join(csv_directory, "REGISTRATION QUERY 3.csv")

        ib_state_tax = os.path.join(csv_directory, "IB STATE TAX.csv")
        corp_state_tax = os.path.join(csv_directory, "CORP STATE TAX.csv")

        wb = openpyxl.Workbook()

        try:
            # ----------------- 1. MB REPORT -----------------
            logger.info("Generating MB-Transactions Sheet...")
            ws_mb = wb.active
            ws_mb.title = "MB-Transactions"
            mb_report.mb_generate_report(ws_mb, report_date)

            mb_daily_data = formatting_data(mb_daily_path)
            mb_monthly_data = formatting_data(mb_monthly_path)

            sql_data_mb = matching(mb_mapper_excel, mb_daily_data, mb_monthly_data)
            mb_final_data = excel_mappers(mb_data, mb_mapper_excel)
            write_mapped_dict_to_excel(ws_mb, mb_final_data, sql_data_mb)

            # ----------------- 2. IB REPORT -----------------
            logger.info("Generating IB-Transactions Sheet...")
            ws_ib = wb.create_sheet(title="IB-Transactions")
            ib_report.ib_generate_report(ws_ib, report_date)

            ib_daily_data = formatting_data(ib_daily_path)
            ib_monthly_data = formatting_data(ib_monthly_path)

            sql_data_ib = matching(ib_mapper_excel, ib_daily_data, ib_monthly_data)
            ib_final_data = excel_mappers(ib_data, ib_mapper_excel)
            write_mapped_dict_to_excel(ws_ib, ib_final_data, sql_data_ib)

            # ----------------- 3. CORPORATE REPORT -----------------
            logger.info("Generating Corporate-Transactions Sheet...")
            ws_corp = wb.create_sheet(title="Corporate-Transactions")
            corp_transactions.corp_generate_report(ws_corp, report_date)

            corp_daily_data = formatting_data(corp_daily_path)
            corp_monthly_data = formatting_data(corp_monthly_path)

            sql_data_corp = matching(corp_mapper_excel, corp_daily_data, corp_monthly_data)
            corp_final_data = excel_mappers(corp_data, corp_mapper_excel)
            corp_write_mapped_dict_to_excel(ws_corp, corp_final_data, sql_data_corp)

            # ----------------- 4. REGISTRATION REPORT -----------------
            logger.info("Generating Registration Sheet...")
            ws_reg = wb.create_sheet(title='Registration')
            reg_report(ws_reg, count_data, corp_ib_count, status_count)

            # ----------------- 5. STATE TAX REPORT -----------------
            logger.info("Applying State Tax Metrics...")
            state_tax(65, ws_ib, ib_state_tax)
            state_tax(18, ws_corp, corp_state_tax)

            # ----------------- ATOMIC SAVE -----------------
            logger.info("Saving compiled Excel report...")
            print(f"\n🎉 Report Successfully Generated: {final_excel_path}\n")
            break

        except Exception as e:
            logger.error(f"Failed during Excel report generation: {e}", exc_info=True)
            print(f"\n❌ Error generating report: {e}\nCheck logs for details.")
            break
        finally:
            wb.save(final_excel_path)
            wb.close()





if __name__ == "__main__":
    report_generation()











