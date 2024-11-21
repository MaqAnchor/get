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
