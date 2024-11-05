
import os
import sys
from datetime import datetime
import pandas as pd


def convert_excel_to_csv_batch(input_folder, separator=";"):
    # Check if input folder exists
    if not os.path.isdir(input_folder):
        print(f"Error: The directory '{input_folder}' does not exist.")
        return

    # Initialize an empty DataFrame to hold all data
    combined_df = pd.DataFrame()

    # Loop through all Excel files in the input directory
    for filename in os.listdir(input_folder):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            # Full path of the input Excel file
            input_file_path = os.path.join(input_folder, filename)

            try:
                # Read the Excel file and append it to the combined DataFrame
                df = pd.read_excel(input_file_path, header=None, engine="openpyxl", dtype=str)
                df.dropna(axis=1, how='all', inplace=True)  # Remove empty columns
                combined_df = pd.concat([combined_df, df], ignore_index=True)
                print(f"Added '{filename}' to the combined DataFrame")
            except Exception as e:
                print(f"Error reading '{filename}': {e}")

    # Generate output file name based on current timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H_%M_%S")
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, f"{timestamp}.csv")

    # Deduplicate by the first column, taking the longest values for the rest of the columns
    deduplicated_df = combined_df.groupby(0).agg(lambda x: max(x, key=lambda v: len(str(v)))).reset_index()

    # Write the deduplicated DataFrame to a single CSV file
    try:
        deduplicated_df.to_csv(output_file, index=False, sep=separator, encoding="utf-8", header=False)
        print(f"Combined CSV saved as '{output_file}'")
    except Exception as e:
        print(f"Error saving combined CSV: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: excel_to_csv_batch <input_folder> [separator]")
    else:
        input_folder = sys.argv[1]
        separator = sys.argv[2] if len(sys.argv) > 2 else ","  # Default to comma
        convert_excel_to_csv_batch(input_folder, separator)
