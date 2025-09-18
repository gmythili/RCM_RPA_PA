import os
import sys
import json
import pandas as pd
from datetime import datetime
import re

skip_markets = ["UL", "UM", "NE", "UG", "MG", "CW", "Other"]

# === Extract market code ===
def get_market_from_cc(cc):
    if len(cc) >= 2 and cc[0].isalpha(): 
        # Check if cc matches MSxxx pattern (MS followed by numbers)
        if re.match(r'^MS\d+', cc, re.IGNORECASE):
            return "MSO"
        return cc[:2]  # First two characters as market code
    return "Other"  # Default if not found

# === Argument: JSON file path ===
if len(sys.argv) != 2:
    print("Usage: python json_to_excel_by_market.py <json_file_path>")
    sys.exit(1)

json_file = sys.argv[1]
if not os.path.isfile(json_file):
    print(f"JSON file not found: {json_file}")
    sys.exit(1)

# === Load JSON data ===
try:
    with open(json_file, 'r', encoding='utf-8') as f:
        json_data = json.load(f)
except (json.JSONDecodeError, FileNotFoundError, OSError) as e:
    print(f"Error reading JSON file: {e}")
    sys.exit(1)

if isinstance(json_data, list):
    records = json_data  # If json_data is a list, use it directly
elif isinstance(json_data, dict):
    records = json_data.get("records", [])
else:
    print("Unexpected JSON structure.")
    sys.exit(1)

print(f"Loaded {len(records)} records from JSON.")

# === Generate market list dynamically ===
markets = set()  # Use a set to avoid duplicates
for record in records:
    try:
        if not isinstance(record, dict):
            continue
        description = record.get("Description", "")
        if not description or not isinstance(description, str):
            continue
        
        if "Care Center:" in description and "Job Title:" in description:
            care_center_part = description.split("Care Center:")[1].split("Job Title:")[0].strip()
            if care_center_part:
                care_center_words = care_center_part.split()
                if care_center_words:
                    care_center = care_center_words[0]
                    market = get_market_from_cc(care_center)
                    if market and market != "Other":
                        markets.add(market)
    except (IndexError, AttributeError, TypeError):
        continue

markets = sorted(markets)  # Convert to a sorted list
market_data = {m: [] for m in markets}
market_data["Other"] = []  # Add a catch-all for unknown/missing CC

# === Extraction function ===
def extract_info(record):
    # Handle case where record might not be a dictionary
    if not isinstance(record, dict):
        print(f"Warning: Invalid record type encountered: {type(record)}")
        return {
            "FN": "Unknown", "LN": "Unknown", "EMAIL": "", "CC": "", "Market": "Other",
            "SF_CaseNumber": "", "SF_ID": "", "SF_CreatedDate": ""
        }
    
    subject = record.get("Subject", "")
    description = record.get("Description", "")

    # Name extraction from subject
    fn = ln = "Unknown"
    if subject and " - " in subject:
        try:
            name_part = subject.split(" - ")[-1].strip()
            if name_part:  # Check if name_part is not empty
                name_parts = name_part.split(",")
                if len(name_parts) == 2:
                    ln_raw = name_parts[0].strip()
                    fn_raw = name_parts[1].strip()
                    if ln_raw:  # Only capitalize if not empty
                        ln = ln_raw.capitalize()
                    if fn_raw:  # Only process if not empty
                        fn_parts = fn_raw.split()
                        if fn_parts:  # Check if split result is not empty
                            fn = fn_parts[0].capitalize()
        except (IndexError, AttributeError, TypeError) as e:
            fn = ln = "Unknown"

    # Email extraction
    email = ""
    if description:
        try:
            if "Email:" in description and "Entity:" in description:
                email_part = description.split("Email:")[1].split("Entity:")[0].strip()
                if email_part:  # Only assign if not empty
                    email = email_part
        except (IndexError, AttributeError, TypeError):
            email = ""

    # Care Center extraction
    cc = ""
    market = "Other"
    if description:
        try:
            if "Care Center:" in description and "Job Title:" in description:
                cc_part = description.split("Care Center:")[1].split("Job Title:")[0].strip()
                if cc_part:  # Check if not empty
                    cc_words = cc_part.split()
                    if cc_words:  # Check if split result is not empty
                        care_center = cc_words[0]
                        cc = care_center
                        market = get_market_from_cc(cc)
        except (IndexError, AttributeError, TypeError):
            pass  # cc and market already set to defaults

    # CreatedDate formatting
    created_date_formatted = ""
    created_date_raw = record.get("CreatedDate", "")
    if created_date_raw and isinstance(created_date_raw, str):
        try:
            # Handle different datetime formats
            clean_date = created_date_raw.replace("Z", "+00:00")
            dt = datetime.fromisoformat(clean_date)
            created_date_formatted = dt.strftime("%m-%d-%Y")
        except (ValueError, TypeError, AttributeError):
            try:
                # Try alternative parsing if fromisoformat fails
                from dateutil import parser
                dt = parser.parse(created_date_raw)
                created_date_formatted = dt.strftime("%m-%d-%Y")
            except:
                created_date_formatted = ""

    # Safe case number processing
    case_number = ""
    case_raw = record.get("CaseNumber", "")
    if case_raw and isinstance(case_raw, str):
        try:
            case_number = case_raw.lstrip("0")
        except (AttributeError, TypeError):
            case_number = str(case_raw)

    return {
        "FN": fn,
        "LN": ln,
        "EMAIL": email,
        "CC": cc,
        "Market": market,
        "SF_CaseNumber": case_number,
        "SF_ID": str(record.get("Id", "")),
        "SF_CreatedDate": created_date_formatted
    }

# === Process records ===
processed_count = 0
error_count = 0

for record in records:
    try:
        info = extract_info(record)
        if not info:  # Skip if extraction failed completely
            error_count += 1
            continue
            
        sheet_name = info["Market"] if info["Market"] in markets else "Other"
        
        # Validate sheet_name is safe for Excel
        if not sheet_name or len(sheet_name) > 31:  # Excel sheet name limit
            sheet_name = "Other"
        
        market_data[sheet_name].append([
            info["FN"], info["LN"], info["EMAIL"], info["CC"], info["Market"], 
            info["SF_CaseNumber"], info["SF_ID"], info["SF_CreatedDate"]
        ])
        processed_count += 1
        
    except Exception as e:
        print(f"Warning: Error processing record: {e}")
        error_count += 1
        continue

print(f"Successfully processed {processed_count} records.")
if error_count > 0:
    print(f"Encountered errors in {error_count} records.")

# === Output to Excel ===
output_date = datetime.now().strftime('%d%m%Y')
output_name = f"UserAccountDeactivationReport_{output_date}.xlsx"
output_folder = r"C:\RPA\PortalTerminationDevelopment\UserExportFile\SF" 
output_path = os.path.join(output_folder, output_name)
# === Ensure output directory exists and is writable ===
try:
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")
    
    # Verify folder was actually created and is accessible
    if not os.path.exists(output_folder):
        print(f"Error: Output folder could not be created or accessed: {output_folder}")
        sys.exit(1)
    
    # Test write permissions by creating a temporary file
    test_file = os.path.join(output_folder, "temp_test_file.tmp")
    try:
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
    except (OSError, IOError, PermissionError) as e:
        print(f"Error: No write permission to output folder {output_folder}: {e}")
        sys.exit(1)
        
except (OSError, IOError, PermissionError) as e:
    print(f"Error: Failed to create or access output folder {output_folder}: {e}")
    sys.exit(1)

# === Write Excel file with comprehensive error handling ===
try:
    # Check if any data exists to write
    total_records = sum(len(rows) for rows in market_data.values())
    if total_records == 0:
        print("Warning: No records found to write to Excel file.")
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheets_created = 0
        for market, rows in market_data.items():
            if market in skip_markets: #Skips writing the markets present in Skip_Markets.
                pass
            else: 
                if rows:
                    try:
                        df = pd.DataFrame(rows, columns=["FN", "LN", "EMAIL", "CC", "Market", "SF_CaseNumber", "SF_ID", "SF_CreatedDate"])
                        df.to_excel(writer, sheet_name=market, index=False)
                        sheets_created += 1
                    except Exception as e:
                        print(f"Error creating sheet '{market}': {e}")
                        continue
        
        if sheets_created == 0:
            print("Error: No sheets could be created in Excel file.")
            sys.exit(1)
            
except (OSError, IOError, PermissionError, ImportError) as e:
    print(f"Error: Failed to create Excel file {output_path}: {e}")
    sys.exit(1)
except Exception as e:
    print(f"Unexpected error while writing Excel file: {e}")
    sys.exit(1)

# === Verify file was actually created ===
try:
    if not os.path.exists(output_path):
        print(f"Error: Excel file was not created at expected location: {output_path}")
        sys.exit(1)
    
    # Check if file has content (not empty)
    file_size = os.path.getsize(output_path)
    if file_size == 0:
        print(f"Error: Excel file was created but is empty: {output_path}")
        sys.exit(1)
        
    print(f"Excel file successfully saved at: {output_path}")
    
except (OSError, IOError) as e:
    print(f"Error verifying Excel file creation: {e}")
    sys.exit(1)