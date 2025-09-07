import json
import sys
import os
from openpyxl import Workbook

def json_to_excel_all(json_file):
    # Load JSON file
    with open(json_file, "r") as f:
        data = json.load(f)

    results = data.get("results", [])

    # Generate Excel filename
    base_name = os.path.splitext(os.path.basename(json_file))[0]
    excel_file = f"{base_name}_all.xlsx"

    # Create workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Policies"

    # Write headers
    headers = ["id", "name", "url-access-policy", "allow-method-policy"]
    ws.append(headers)

    # Write all rows
    for item in results:
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
        print("Usage: python json_to_excel_all.py <input.json>")
        sys.exit(1)

    input_json = sys.argv[1]
    json_to_excel_all(input_json)
