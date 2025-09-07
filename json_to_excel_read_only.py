import json
import sys
import os
from openpyxl import Workbook

def json_to_excel_non_empty(json_file):
    # Load JSON file
    with open(json_file, "r") as f:
        data = json.load(f)

    results = data.get("results", [])

    # Generate Excel filename
    base_name = os.path.splitext(os.path.basename(json_file))[0]
    excel_file = f"{base_name}_non_empty.xlsx"

    # Create workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Filtered Policies"

    # Write headers
    headers = ["id", "name", "url-access-policy", "allow-method-policy"]
    ws.append(headers)

    # Write rows where "allow-method-policy" is not empty
    for item in results:
        if item.get("allow-method-policy", "") != "":
            row = [
                item.get("id", ""),
                item.get("name", ""),
                item.get("url-access-policy", ""),
                item.get("allow-method-policy", "")
            ]
            ws.append(row)

    # Save Excel
    wb.save(excel_file)
    print(f"✅ Excel file created: {excel_file}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <input.json>")
        sys.exit(1)

    input_json = sys.argv[1]
    json_to_excel_non_empty(input_json)

