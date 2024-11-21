import os
import pandas as pd

# Define the filename patterns
filename_patterns = [
    'QDS-above-70-crossed-40d',
    'QDS-0-69-crossed-40d',
    'QDS-0-69-less-40d',
    'QDS-above-70-less-40d'
]

# Function to process each file
def process_file(file_path):
    try:
        print(f"Processing file: {file_path}")
        df = pd.read_csv(file_path)

        # Check if 'Last check-in Intune' column exists
        if 'Last check-in Intune' in df.columns:
            # Replace 'Blank' with '01-01-1900' and 'Not Found' with '01-01-1800'
            df['Last check-in Intune'] = df['Last check-in Intune'].replace(
                {'Blank': '01-01-1900', 'Not Found': '01-01-1800'}
            )

            # Save the updated CSV file, retaining the original filename
            df.to_csv(file_path, index=False)
            print(f"File successfully updated: {file_path}")
        else:
            print(f"'Last check-in Intune' column not found in {file_path}. Skipping file.")

    except Exception as e:
        print(f"An error occurred while processing {file_path}: {e}")

# Main script
if __name__ == "__main__":
    print("Starting script to process CSV files...")
    files_found = False

    # Loop through the filename patterns
    for pattern in filename_patterns:
        # Find matching files in the current directory
        for file in os.listdir('.'):
            if pattern in file and file.endswith('.csv'):
                files_found = True
                process_file(file)
    
    if not files_found:
        print("No matching files found. Exiting script.")
    else:
        print("All matching files processed.")


















































import os
import pandas as pd
from datetime import datetime
import time

# Start the timer
start_time = time.time()

# Define filename patterns to look for
filename_patterns = [
    'QDS-above-70-crossed-40d',
    'QDS-0-69-crossed-40d',
    'QDS-0-69-less-40d',
    'QDS-above-70-less-40d'
]

# Get the current working directory
cwd = os.getcwd()

# Get the list of files in the current directory
files = os.listdir(cwd)

# Filter files matching the patterns
csv_files = [f for f in files if any(pattern in f for pattern in filename_patterns) and f.endswith('.csv')]

# Initialize an empty list to hold dataframes
dataframes = []

# Process each file
print("Starting the merging process...\n")
for file in csv_files:
    print(f"Processing file: {file}")
    try:
        # Read the CSV file
        df = pd.read_csv(file, dtype=str)  # Read as strings to handle mixed-type data
        df['origin'] = file  # Add the 'origin' column
        dataframes.append(df)
        print(f"Successfully processed {file}. Rows added: {len(df)}\n")
    except Exception as e:
        print(f"Failed to process {file}. Error: {e}\n")

# Combine all dataframes
if dataframes:
    print("Merging all files...")
    merged_df = pd.concat(dataframes, ignore_index=True)

    # Get today's date in the specified format
    today_date = datetime.now().strftime('%d-%b-%y')

    # Define the output file name
    output_file = f"{today_date}.csv"

    # Save the merged data to a new CSV file
    merged_df.to_csv(output_file, index=False)

    print(f"\nMerging complete! Merged file saved as: {output_file}")
else:
    print("No files matched the patterns. Exiting without creating a merged file.")

# Calculate and display the elapsed time
end_time = time.time()
elapsed_time = end_time - start_time
print(f"Total time taken: {elapsed_time:.2f} seconds.")









