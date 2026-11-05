import os
import csv
import logging
import psycopg2
from psycopg2 import Error
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("RawQueryExtractor")

# ==========================================
# CONFIGURATION
# ==========================================


# Directly writing the absolute path
env_path = '/home/yashraj/reports/.env'

if os.path.exists(env_path):
    with open(env_path, 'r') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                os.environ[key.strip()] = value.strip()

POSTGRES_CONFIG = {
    'host': os.getenv('DB_HOST'),
    'port': os.getenv('DB_PORT'),
    'database': os.getenv('DB_NAME'),
    'user': os.getenv('DBUSER'),
    'password': os.getenv('DBPASS'),
    'sslmode': 'require'
}



# List your sheet names and their corresponding SQL files
NIGHT_REPORT_PIPELINE_STEPS = {
    "IB QUERY DAILY" :{"sql_file": "night_ib_query_daily.txt"},
    "IB MONTHLY QUERY":{"sql_file": "night_ib_query_monthly.txt"},
    "IB STATE TAX":{"sql_file": "night_ib_state_tax.txt"},
    "IB TIN AND GST":{"sql_file": "night_ib_tin_and_gst.txt"},
    "MB QUERY DAILY":{"sql_file": "night_mb_query_daily.txt"},
    "MB MONTHLY QUERY":{"sql_file": "night_mb_query_monthly.txt"},
    "CORP QUERY DAILY":{"sql_file": "night_corp_query_daily.txt"},
    "CORP QUERY MONTHLY":{"sql_file": "night_corp_query_monthly.txt"},
    "CORP STATE TAX":{"sql_file": "night_corp_state_tax.txt"},
    "REGISTRATION QUERY 1":{"sql_file": "night_registration_query_1.txt"},
    "REGISTRATION QUERY 2":{"sql_file": "night_registration_query_2.txt"},
    "REGISTRATION QUERY 3":{"sql_file": "night_registration_query_3.txt"},
}

DAY_REPORT_PIPELINE_STEPS = {
    "IB QUERY DAILY" :{"sql_file": "night_ib_query_daily.txt"},
    "IB MONTHLY QUERY":{"sql_file": "night_ib_query_monthly.txt"},
    "IB STATE TAX":{"sql_file": "night_ib_state_tax.txt"},
    "IB TIN AND GST":{"sql_file": "night_ib_tin_and_gst.txt"},
    "MB QUERY DAILY":{"sql_file": "night_mb_query_daily.txt"},
    "MB MONTHLY QUERY":{"sql_file": "night_mb_query_monthly.txt"},
    "CORP QUERY DAILY":{"sql_file": "night_corp_query_daily.txt"},
    "CORP QUERY MONTHLY":{"sql_file": "night_corp_query_monthly.txt"},
    "CORP STATE TAX":{"sql_file": "night_corp_state_tax.txt"},
    "REGISTRATION QUERY 1":{"sql_file": "night_registration_query_1.txt"},
    "REGISTRATION QUERY 2":{"sql_file": "night_registration_query_2.txt"},
    "REGISTRATION QUERY 3":{"sql_file": "night_registration_query_3.txt"},
}






#load_sql is used to extract to read the sql file present in the  particular file_path


def load_sql(file_path_queries,filename: str):
    filepath = os.path.join(file_path_queries, filename)
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"SQL file not found at: {filepath}")
    with open(filepath, "r", encoding="utf-8") as f:
        return f.read().strip()


#Report Timeline are only NIGHT_REPORT_PIPELINE_STEPS  or DAY_REPORT_PIPELINE_STEPS  depending on them there are
#different queries stored in different file_paths

def extract_and_dump_csv(report_timeline,pipeline_type):
    """Executes SQL queries and writes exact result sets directly to CSV files."""
    # os.makedirs(STAGING_CSV_DIR, exist_ok=True)
    conn = None
    cursor = None

    if pipeline_type== 'NIGHT_REPORT_PIPELINE_STEPS':
            sql_dir = os.path.abspath(r"/home/yashraj/reports/night_queries")
            staging_dir= os.path.abspath(r"/home/yashraj/reports/night_staging_csv")
    else:
            sql_dir = os.path.abspath(r"/home/yashraj/reports/day_queries")
            staging_dir = os.path.abspath(r"/home/yashraj/reports/day_staging_csv")


    try:
        logger.info("Connecting to Database...")
        conn = psycopg2.connect(**POSTGRES_CONFIG)
        cursor = conn.cursor()

        for sheet_name, info in report_timeline.items():
            sql_query = load_sql(sql_dir,info["sql_file"])
            logger.info(f"Executing query for '{sheet_name}'...")

            cursor.execute(sql_query)



            # Check if query returned rows
            if cursor.description:
                headers = [col[0] for col in cursor.description]
                target_csv_path = os.path.join(staging_dir, f"{sheet_name}.csv")
                os.makedirs(os.path.dirname(target_csv_path), exist_ok=True)

                # Stream results straight into CSV using built-in csv module (no extra overhead)
                with open(target_csv_path, mode="w", newline="", encoding="utf-8") as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerow(headers)  # Write exact DB column headers

                    row_count = 0
                    while True:
                        rows = cursor.fetchmany(10000)  # Fetch in batches of 10,000 for memory efficiency
                        if not rows:
                            break
                        writer.writerows(rows)
                        row_count += len(rows)

                logger.info(f"Saved {row_count} raw rows to: {target_csv_path}")

        conn.commit()
        logger.info("All query results successfully saved to CSV .")


    except Error as e:
        logger.error(f"Execution Failed! MySQL Error: {e}")


    finally:
        if cursor:
            cursor.close()
        if conn and conn.closed==0:
            conn.close()
            logger.info("Database connection closed.")



def run_extraction_cli():
    print("==================================================")
    print("      STEP 1: DATA EXTRACTION & STAGING MENU      ")
    print("==================================================")
    print("A. For Night-Report data download from DB")
    print("B. For Day-Report data download from DB")


    choice = input("\nSelect Option (A or B): ").strip().upper()

    if choice == "A":
        success = extract_and_dump_csv(NIGHT_REPORT_PIPELINE_STEPS,'NIGHT_REPORT_PIPELINE_STEPS')
        if success:
            print("Now you can successfully run the report_generation.bat file ")
            return True
        else:
            print("Please run queries manually inside the night_queries Folder \n"
                  "Keep name of the csv files same as the .sql file \n"
                  "Place exported CSVs into './night_staging_csv/', and then only run report_generation.py.")
            return False

    elif choice == "B":
        success = extract_and_dump_csv(DAY_REPORT_PIPELINE_STEPS,'DAY_REPORT_PIPELINE_STEPS')
        if success:
            print("Now you can successfully run the report_generation.bat file ")
            return True
        else:
            print("Please run queries manually inside the day_queries Folder \n"
                  "Keep name of the csv files same as the .sql file \n"
                  "Place exported CSVs into './day_staging_csv/', and then only run report_generation.bat")
            return False
    else:
        print("\n Invalid choice. Exiting.")
        return False


if __name__ == "__main__":
    run_extraction_cli()