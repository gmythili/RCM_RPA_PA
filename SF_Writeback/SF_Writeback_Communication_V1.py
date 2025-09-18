import pandas as pd
import sys
import os
import time
from datetime import datetime

# Markets to skip
skip_markets = ["UL", "UM", "NE", "UG", "CW", "Other"]

# Columns to skip (partial match allowed)
skip_columns = ["FN", "LN", "EMAIL", "CC","Market", "SF_CaseNumber", "SF_CreatedDate", "CaseNumber"]

def wait_for_file(filepath, timeout=20):
    print(f"Waiting for file: {filepath}")
    if not os.path.exists(filepath):
        print(f"Error: File not found -> {filepath}")
        sys.exit(1)

    if os.path.basename(filepath).startswith('~$'):
        print(f"Error: The file {filepath} looks like a temporary Excel lock file.")
        sys.exit(1)

    start_time = time.time()
    while True:
        try:
            with open(filepath, 'rb'):
                print(f"File ready: {filepath}")
                return
        except Exception as e:
            print(f"File access Error: {e}")
        if time.time() - start_time > timeout:
            print(f"Timeout: File not accessible after {timeout} seconds -> {filepath}")
            sys.exit(1)
        time.sleep(1)

# === Startup ===
time.sleep(1)

# === Get command line args ===
if len(sys.argv) != 2:
    print("Usage: python SF_Writeback_Sahara.py <report_excel_path>")
    sys.exit(1)

report_path = sys.argv[1]
wait_for_file(report_path)

try:
    all_data = []

    # Load Excel file
    try:
        xls = pd.ExcelFile(report_path)
    except Exception as e:
        print(f"Error: Unable to open Excel file -> {e}")
        sys.exit(1)

    for sheet_name in xls.sheet_names:
        if "report" in sheet_name.lower():
            continue
        if sheet_name in skip_markets:
            continue

        try:
            df = pd.read_excel(report_path, sheet_name=sheet_name, engine="openpyxl")
        except Exception as e:
            print(f"Error reading sheet '{sheet_name}': {e}")
            continue

        # Find SF_ID column dynamically using contains "ID"
        sf_id_col = None
        for col in df.columns:
            if "id" in str(col).lower():  
                sf_id_col = col
                break

        if not sf_id_col:
            print(f"Warning: Skipping sheet '{sheet_name}' — no column containing 'ID' found")
            continue

        # Determine dynamic columns (skip unwanted ones using contains)
        dynamic_cols = [
            col for col in df.columns
            if not any(skip_key.lower() in str(col).lower() for skip_key in skip_columns)
            and col != sf_id_col
        ]

        if not dynamic_cols:
            print(f"Warning: Skipping sheet '{sheet_name}' — no valid dynamic columns found")
            continue

        # Process rows
        for _, row in df.iterrows():
            try:
                sf_id = str(row[sf_id_col]).strip()
                if not sf_id or sf_id.lower() == "nan":
                    continue  # skip empty IDs

                comm_parts = []
                for col in dynamic_cols:
                    try:
                        val = str(row[col]).strip()
                        if val and val.lower() != "nan":
                            comm_parts.append(f"{col} - {val}")
                    except Exception as e:
                        print(f"Warning: Error reading column '{col}' in sheet '{sheet_name}': {e}")
                        continue

                if comm_parts:  # only add if we have communication values
                    communication = ",".join(comm_parts)
                    all_data.append({"SF_ID": sf_id, "Communication": communication})

            except Exception as e:
                print(f"Warning: Error processing row in sheet '{sheet_name}': {e}")
                continue

    # Create file name with current date
    today_str = datetime.now().strftime("%Y-%m-%d")
    output_filename = f"{today_str}_SF_Writeback_communication.csv"

    # Dummy new location (replace in production)
    output_csv = os.path.join(r"C:\Users\RCM_RPAdmin\Unified\PAD-UWH-PortalTermination - SF_Reports", output_filename)

    if all_data:  # Only create CSV if we have data
        try:
            pd.DataFrame(all_data).to_csv(output_csv, index=False, lineterminator="\n", encoding="utf-8-sig")
            print(f"Communication CSV created successfully: {output_csv}")
        except Exception as e:
            print(f"Error writing CSV file: {e}")
            sys.exit(1)
    else:
        print("No valid data found — CSV will not be created.")

except Exception as e:
    print("An unexpected Error occurred:", str(e))
    sys.exit(1)
