import sys
import pandas as pd
import re

# Ensure a file name is provided
if len(sys.argv) < 2:
    print("Usage: python regex.py <file.xlsx>")
    sys.exit(1)

# Get the filename from the command-line argument
file_name = sys.argv[1]

# Define sheet name and regex pattern
sheet_name = "Sheet2"
version_pattern = r'\b\d{1,2}\.\d{1,2}\.\d{1,2}\b'

try:
    # Load the Excel sheet
    df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')

    # Ensure the sheet has at least 3 columns (A, B, C)
    if df.shape[1] < 3:
        print("Error: Sheet2 must have at least three columns (A, B, C).")
        sys.exit(1)

    # Extract version numbers from column B (Index 1) and save to column C (Index 2)
    df.iloc[:, 2] = df.iloc[:, 1].astype(str).apply(lambda x: ", ".join(re.findall(version_pattern, x)))

    # Save the updated file
    output_file = file_name.replace(".xlsx", "_updated.xlsx")
    df.to_excel(output_file, sheet_name=sheet_name, index=False, engine='openpyxl')

    print(f"Processed successfully! Extracted versions are saved in column C of {output_file}")

except Exception as e:
    print(f"Error: {e}")
