import sys
import pandas as pd
import re

# Ensure the user provides a filename
if len(sys.argv) < 2:
    print("Usage: python regex.py <file.xlsx>")
    sys.exit(1)

# Get filename from command-line argument
file_name = sys.argv[1]

# Define the sheet name and column to work on
sheet_name = "Sheet2"
column_to_process = "B"
output_column = "C"

# Regex pattern for extracting versions in xx.xx.xx format
version_pattern = r'\b\d{1,2}\.\d{1,2}\.\d{1,2}\b'

try:
    # Load the Excel file
    df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')

    # Ensure column B exists
    if column_to_process not in df.columns:
        print(f"Error: Column {column_to_process} not found in sheet {sheet_name}.")
        sys.exit(1)

    # Extract version numbers from column B
    df[output_column] = df[column_to_process].astype(str).apply(lambda x: ", ".join(re.findall(version_pattern, x)))

    # Save the updated file
    output_file = file_name.replace(".xlsx", "_updated.xlsx")
    df.to_excel(output_file, sheet_name=sheet_name, index=False, engine='openpyxl')

    print(f"Processed successfully! Extracted versions are saved in column {output_column} of {output_file}")

except Exception as e:
    print(f"Error: {e}")
