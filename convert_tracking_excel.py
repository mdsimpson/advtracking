import pandas as pd
import csv
import os
from pathlib import Path

def process_excel_file(file_path):
    xl = pd.ExcelFile(file_path)
    sheet_name = xl.sheet_names[0]
    df = pd.read_excel(file_path, header=None, sheet_name=0)

    company_row = df.iloc[7, 2:]
    companies = company_row[company_row.notna()].tolist()
    result = {company: {} for company in companies}

    pass_metrics = ["Unique Passes in Use", "Contracted", "Activated"]
    for i in range(len(df)):
        cell_value = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""
        if cell_value in pass_metrics:
            values = df.iloc[i, 2:2 + len(companies)]
            for company, value in zip(companies, values):
                if pd.notna(value):
                    result[company][f"Pass:{cell_value}"] = value

    def process_section(start_row, num_rows, prefix):
        for i in range(start_row + 1, start_row + 1 + num_rows):
            title = df.iloc[i, 0]
            if pd.isna(title):
                continue
            values = df.iloc[i, 2:2 + len(companies)]
            for company, value in zip(companies, values):
                if pd.notna(value):
                    key = f"{prefix}:{title.strip()}"
                    result[company][key] = value

    for i in range(len(df)):
        cell_value = str(df.iloc[i, 0]).lower() if pd.notna(df.iloc[i, 0]) else ""
        if cell_value == "transit":
            process_section(i, 4, "Transit")
        elif cell_value == "regional rail":
            process_section(i, 2, "Regional Rail")
        elif cell_value == "paratransit":
            process_section(i, 2, "Paratransit")

    return result, sheet_name


def export_to_csv(data, output_file):
    mapping = [
        ("Unique Card Taps", "Pass:Unique Passes in Use"),
        ("Eligible Employees", "Pass:Contracted"),
        ("Enrolled Employees", "Pass:Activated"),
        ("Total Transit Trips", "Transit:Total Trips"),
        ("Unique Card Taps (Transit)", "Transit:Unique Passes"),
        ("Total Railroad Trips", "Regional Rail:Base Trips"),
        ("Unique Card Taps (Rail)", "Regional Rail:Unique Passes"),
        ("Total CCT Trips", "Paratransit:Base Trips"),
        ("Unique Card Taps (CCT)", "Paratransit:Unique Passes")
    ]

    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['key', 'value'])

        for company in data:
            for output_name, data_key in mapping:
                if data_key in data[company]:
                    writer.writerow([f"{company}:{output_name}", data[company][data_key]])


def main():
    default_path = "input.xlsx"

    # Prompt the user for a file path
    file_path = input(f"Enter the file path (Press Enter to use default: {default_path}): ").strip().strip('"')

    # Use the default path if no input is provided
    if not file_path:
        file_path = default_path

    try:
        data, sheet_name = process_excel_file(file_path)
        # Printing the contents of the dictionary of dictionaries
        for key, inner_dict in data.items():
            print(f"{key}:")
            for inner_key, value in inner_dict.items():
                print(f"  {inner_key}: {value}")
            print()  # Add a blank line for better readability

        output_csv = f"{sheet_name}.csv"
        export_to_csv(data, output_csv)
        print(f'Data exported to "{Path(output_csv).absolute()}"')
    except Exception as e:
        print(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    main()