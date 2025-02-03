import json
import pandas as pd
import ijson  # For efficient JSON streaming

# Input file paths
json_file_path = "input.json"
xlsx_file_path = "input.xlsx"
output_file_path = "output.xlsx"

# Load JSON data using a streaming approach
def load_json(file_path):
    records = []
    with open(file_path, "r") as f:
        for record in ijson.items(f, "item"):
            records.append(record)
    return records

# Load Excel data and convert to a set of serial numbers
def load_excel(file_path):
    df = pd.read_excel(file_path)
    return set(df["Combined ID Serial Number"].astype(str))  # Ensure all are strings

# Find unmatched records
def find_unmatched_records(json_data, excel_serial_numbers):
    unmatched_records = []
    for record in json_data:
        device_id = record.get("DeviceId")
        if device_id and str(device_id) not in excel_serial_numbers:  # Compare as strings
            # Include only the desired fields
            unmatched_records.append({
                "DeviceId": record.get("DeviceId"),
                "ComputerName": record.get("ComputerName"),
                "USBDevice": record.get("USBDevice"),
            })
    return unmatched_records

# Save unmatched records to Excel in batches if necessary
def save_excel(data, file_path, batch_size=1000000):
    if len(data) <= batch_size:
        pd.DataFrame(data).to_excel(file_path, index=False)  # Save as Excel without the index column
    else:
        for i in range(0, len(data), batch_size):
            batch = data[i:i + batch_size]
            batch_file_path = f"{file_path.rstrip('.xlsx')}_part{i // batch_size + 1}.xlsx"
            pd.DataFrame(batch).to_excel(batch_file_path, index=False)
            print(f"Batch saved to {batch_file_path}")

# Main execution
if __name__ == "__main__":
    # Load data
    print("Loading JSON data...")
    json_data = load_json(json_file_path)
    
    print("Loading Excel data...")
    excel_serial_numbers = load_excel(xlsx_file_path)

    # Find unmatched records
    print("Finding unmatched records...")
    unmatched_records = find_unmatched_records(json_data, excel_serial_numbers)

    # Save unmatched records to output Excel file
    if unmatched_records:
        print(f"Saving {len(unmatched_records)} unmatched records to Excel...")
        save_excel(unmatched_records, output_file_path)
        print(f"Unmatched records saved to {output_file_path}")
    else:
        print("No unmatched records found.")
