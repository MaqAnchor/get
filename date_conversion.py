import os
import glob
import csv
import re
import datetime

# Define filename patterns to look for
filename_patterns = [
    'QDS-above-70-crossed-40d',
    'QDS-0-69-crossed-40d',
    'QDS-0-69-less-40d',
    'QDS-above-70-less-40d'
]

# Regex to match date in dd-mm-yyyy (or d-m-yyyy) format
date_pattern = re.compile(r'^(\d{1,2})-(\d{1,2})-(\d{4})$')

def convert_date_format(value):
    """
    If 'value' looks like a date in dd-mm-yyyy, convert to mm-dd-yyyy.
    Otherwise return it unchanged.
    """
    match = date_pattern.match(value.strip())
    if match:
        day, month, year = match.groups()
        try:
            # Convert to datetime to ensure it's a valid date
            dt = datetime.datetime.strptime(value, '%d-%m-%Y')
            return dt.strftime('%m-%d-%Y')
        except ValueError:
            # If it fails to parse, just return original
            return value
    else:
        return value

for pattern in filename_patterns:
    # For each pattern, find all CSV files that match
    for file_path in glob.glob(pattern + '.csv'):
        print(f'Processing file: {file_path}')
        
        # Read the entire CSV
        with open(file_path, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            rows = list(reader)
        
        # Convert dates and store updated rows
        updated_rows = []
        for row in rows:
            new_row = [convert_date_format(cell) for cell in row]
            updated_rows.append(new_row)

        # Option 1: Overwrite existing file
        with open(file_path, 'w', newline='', encoding='utf-8') as outfile:
            writer = csv.writer(outfile)
            writer.writerows(updated_rows)

        # Option 2: Write to a new file (uncomment if preferred)
        # new_file_path = file_path.replace('.csv', '_updated.csv')
        # with open(new_file_path, 'w', newline='', encoding='utf-8') as outfile:
        #     writer = csv.writer(outfile)
        #     writer.writerows(updated_rows)
        # print(f'Updated file saved as: {new_file_path}')
