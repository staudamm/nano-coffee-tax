import argparse
from datetime import datetime, timedelta
import excel
from openpyxl import load_workbook, Workbook
import pandas as pd
import os
import sys

TEMPLATE_EXCEL_FILE_PATH = "../NANO_KAFFEE_GmbH_YYYY_MM_Abt.3A.xlsx"


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
        self.eu_amount = 0

    def _populate_row(self, raw_data, idx):
        row = excel.row.copy()
        for key, source_idx in excel.source_idx_from_key.items():
            row[key] = raw_data[source_idx]/1000 if key == "Amount" else raw_data[source_idx]
        row["ID"] = idx + 1
        self.eu_amount += row["Amount"]
        self.ws.append(list(row.values()))

    def append_csv_to_xlsx(self, source_df: pd.DataFrame):
        # rows_to_keep = list(self.ws.iter_rows(min_row=1, max_row=excel.HEADER_ROW, values_only=True))
        self.ws.delete_rows(excel.HEADER_ROW+1, self.ws.max_row)

        # Append the CSV data to the XLSX file
        for idx, raw_data in enumerate(source_df.itertuples(index=False, name=None)):
            self._populate_row(raw_data, idx)

    def save(self, target_path=TEMPLATE_EXCEL_FILE_PATH):
        self.ws[excel.TOTAL_EU] = self.eu_amount
        start_date, end_date = get_previous_month_range()
        self.ws[excel.TIME_FROM] = start_date.strftime('%d.%m.%Y')
        self.ws[excel.TIME_TO] = end_date.strftime('%d.%m.%Y')
        target_path = target_path.replace("YYYY", start_date.strftime('%Y'))
        target_path = target_path.replace("MM", start_date.strftime('%m'))
        print
        # Save the updated XLSX file
        self.wb.save(target_path)
        # print(f"Appended CSV content to '{xlsx_file}', keeping the first 10 rows.")


def main():
    # Create an argument parser
    parser = argparse.ArgumentParser(description="A script that processes a .csv input file and saves to an output file.")

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

    # Load the input file as a Pandas DataFrame
    try:
        df = pd.read_csv(args.input_file)
        print("Input file loaded successfully.")
        print(df.head())  # Display the first few rows of the DataFrame
    except Exception as e:
        print(f"Error loading input file: {e}")
        sys.exit(1)

    report = A3Report()
    report.append_csv_to_xlsx(df)
    report.save()


    # # Save the DataFrame to the output file
    # try:
    #     df.to_csv(args.output_file, index=False)
    #     print(f"DataFrame saved to '{args.output_file}'.")
    # except Exception as e:
    #     print(f"Error saving output file: {e}")
    #     sys.exit(1)

if __name__ == "__main__":
    main()
