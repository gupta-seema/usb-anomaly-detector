# USB Anomaly Detector

## Overview
The **USB Anomaly Detector** is a Python-based tool designed to identify discrepancies between USB device connection details and a company's official records. It processes a JSON file containing USB device data and an Excel file containing the company's records, then outputs a list of unmatched records for further analysis.

## Features
- **Efficient Data Handling:** Streams JSON data for performance and handles large datasets efficiently.
- **Customizable Output:** Filters and saves unmatched records with only the required fields.
- **Batch Processing:** Splits output into multiple files if unmatched records exceed Excel's row limit.
- **Easy Integration:** Designed for easy use and integration into workflows.

## Use Case
This tool is ideal for IT administrators and analysts who need to:
- Verify USB device connections against authorized company records.
- Detect unauthorized or unregistered USB devices.
- Generate reports for further compliance and security analysis.

## Prerequisites
Ensure you have the following installed:
- **Python 3.7+**
- Required Python libraries:
  - `pandas`
  - `ijson`
  - `openpyxl`

Install dependencies using:
```bash
pip install pandas ijson openpyxl
```

## Input Files
1. **JSON File** (`input.json`):
   - Contains records with fields such as `DeviceId`, `ComputerName`, and `USBDevice`.
2. **Excel File** (`input.xlsx`):
   - Contains a column named `Combined ID Serial Number` with a list of registered device IDs.

## Output
- Generates an Excel file (`output.xlsx`) containing unmatched USB records.
- If unmatched records exceed Excel's row limit (1,048,576 rows), the output is split into multiple files (e.g., `output_part1.xlsx`, `output_part2.xlsx`, etc.).

## Usage
1. Place the input JSON file and Excel file in the same directory as the script.
2. Update the file paths in the script if necessary:
   ```python
   json_file_path = "input.json"
   xlsx_file_path = "input.xlsx"
   output_file_path = "output.xlsx"
   ```
3. Run the script:
   ```bash
   python comparefiles.py
   ```
4. The script will:
   - Load the JSON and Excel data.
   - Compare the `DeviceId` values in the JSON file to the serial numbers in the Excel file.
   - Save the unmatched records to the specified output file.

## How It Works
1. **Load JSON Data:**
   - Streams records from the JSON file using `ijson` for memory efficiency.
2. **Load Excel Data:**
   - Reads the Excel file and extracts the `Combined ID Serial Number` column into a set for fast lookup.
3. **Compare Records:**
   - Checks if each `DeviceId` from the JSON file exists in the company's records.
   - Collects unmatched records, including fields `DeviceId`, `ComputerName`, and `USBDevice`.
4. **Save Output:**
   - Writes unmatched records to an Excel file, splitting into batches if necessary.

## Example
### Input JSON (`input.json`):
```json
[
  {"DeviceId": "12345", "ComputerName": "PC-01", "USBDevice": "USB-A"},
  {"DeviceId": "67890", "ComputerName": "PC-02", "USBDevice": "USB-B"}
]
```

### Input Excel (`input.xlsx`):
| Combined ID Serial Number |
|---------------------------|
| 12345                    |

### Output Excel (`output.xlsx`):
| DeviceId | ComputerName | USBDevice |
|----------|--------------|-----------|
| 67890    | PC-02        | USB-B     |

## Troubleshooting
1. **No Output or Empty Output:**
   - Verify that the column name in the Excel file is `Combined ID Serial Number`.
   - Ensure data types for `DeviceId` and `Combined ID Serial Number` match (both treated as strings).
2. **File Not Found Errors:**
   - Confirm the file paths are correct and files exist in the specified locations.
3. **Performance Issues:**
   - For very large datasets, consider increasing system memory or processing files in smaller chunks.

## Customization
- To include additional fields in the output, update the `find_unmatched_records` function:
   ```python
   unmatched_records.append({
       "DeviceId": record.get("DeviceId"),
       "ComputerName": record.get("ComputerName"),
       "USBDevice": record.get("USBDevice"),
       "OtherField": record.get("OtherField")
   })
   ```
- Adjust batch size for large outputs:
   ```python
   save_excel(data, file_path, batch_size=500000)  # Example: Reduce batch size to 500,000 rows
   ```

## License
This project is open-source and available under the MIT License. Feel free to use, modify, and distribute it.

## Author
Created by [Your Name]. This tool was designed to enhance IT security and compliance workflows by identifying unauthorized USB connections.

