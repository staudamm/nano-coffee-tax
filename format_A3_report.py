import argparse
from datetime import datetime, timedelta
import excel
from openpyxl import load_workbook, Workbook
import csv
import sys
import os

TEMPLATE_EXCEL_FILE_PATH = "NANO_KAFFEE_GmbH_YYYY_MM_Abt.3A.xlsx"


def get_previous_month_range():
    today = datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)

    return first_day_of_previous_month, last_day_of_previous_month


class A3Report:
    def __init__(self):
        self.wb = load_workbook(TEMPLATE_EXCEL_FILE_PATH)
        self.ws = self.wb.active
        self.amount = {"EU": 0, "Ausfuhr": 0}

    def _populate_row(self, raw_data, idx):
        row = excel.row.copy()
        for key, source_idx in excel.source_idx_from_key.items():
            row[key] = int(raw_data[source_idx])/1000 if key == "Amount" else raw_data[source_idx]
        row["ID"] = idx + 1
        row["Region"] = "EU" if "B2B" in raw_data[9] else "Ausfuhr"
        self.amount[row["Region"]] += row["Amount"]
        self.ws.append(list(row.values()))

    def append_csv_to_xlsx(self, file_path):
        self.ws.delete_rows(excel.HEADER_ROW+1, self.ws.max_row)
        # Append the CSV data to the XLSX file
        with open(file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # Skip the first row
            idx = 0
            for raw_data in reader:
                self._populate_row(raw_data, idx)
                idx += 1

    def save(self, target_path=TEMPLATE_EXCEL_FILE_PATH):
        self.ws[excel.AMOUNT["EU"]] = self.amount["EU"]
        self.ws[excel.AMOUNT["Ausfuhr"]] = self.amount["Ausfuhr"]
        start_date, end_date = get_previous_month_range()
        self.ws[excel.TIME_FROM] = start_date.strftime('%d.%m.%Y')
        self.ws[excel.TIME_TO] = end_date.strftime('%d.%m.%Y')
        target_path = target_path.replace("YYYY", start_date.strftime('%Y'))
        target_path = target_path.replace("MM", start_date.strftime('%m'))
        # Save the updated XLSX file
        self.wb.save(target_path)


def main():
    # Create an argument parser
    parser = argparse.ArgumentParser()

    # Add arguments for input and output files
    parser.add_argument("input_file", type=str, help="Path to the input .csv file")
    parser.add_argument("output_file", type=str, help="Path to the output file")

    # Parse the arguments
    args = parser.parse_args()

    # Validate input file
    if not args.input_file.endswith('.csv'):
        print("Error: The input file must be a .csv file.")
        sys.exit(1)

    if not os.path.exists(args.input_file):
        print(f"Error: The input file '{args.input_file}' does not exist.")
        sys.exit(1)

    report = A3Report()
    report.append_csv_to_xlsx(args.input_file)
    report.save()

if __name__ == "__main__":
    main()
