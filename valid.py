import os
import pandas as pd
from datetime import datetime
import logging
from tqdm import tqdm
import time

def main():
    # Record the start time
    start_time = time.time()

    # Configure logging to append to the log file
    log_file = 'script.log'
    logging.basicConfig(
        filename=log_file,
        filemode='a',  # Append mode
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    # Start logging
    logging.info("Script execution started.")

    # Call the functions
    add_ad_column_to_csv_files()
    add_intune_columns_to_csv_files()
    add_arch_and_based_columns()

    # Calculate total execution time
    end_time = time.time()
    elapsed_time = end_time - start_time
    elapsed_time_str = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))

    # Display total execution time
    print(f"\nTotal execution time: {elapsed_time_str}")

    # Log total execution time
    logging.info(f"Total execution time: {elapsed_time_str}")
    logging.info("Script execution completed.\n")

def add_ad_column_to_csv_files():
    print("Starting the process to add 'AD' column to CSV files based on AD_Dump.xlsx data...")

    # Load the AD_Dump.xlsx data
    try:
        print("Loading AD_Dump.xlsx...")
        ad_data = pd.read_excel('AD_Dump.xlsx', usecols=['Computer Name', 'OU'])
        
        # Clean and standardize 'Computer Name' for consistent matching
        ad_data['Computer Name'] = ad_data['Computer Name'].astype(str).str.strip().str.lower()
        
        # Create a lookup dictionary with 'AD' as the column name
        ad_lookup = dict(zip(ad_data['Computer Name'], ad_data['OU']))
        logging.info("Successfully loaded AD_Dump.xlsx and prepared lookup dictionary.")
    except Exception as e:
        logging.error(f"Failed to load AD_Dump.xlsx: {e}")
        print("Error: Could not load AD_Dump.xlsx. Check log for details.")
        return

    # List of CSV files to process
    csv_files = [
        "QDS-above-70-crossed-40d.csv",
        "QDS-0-69-crossed-40d.csv",
        "QDS-0-69-less-40d.csv",
        "QDS-above-70-less-40d.csv"
    ]

    # Process each CSV file with tqdm for a progress bar
    for csv_file in tqdm(csv_files, desc="Processing CSV files for AD data", unit="file"):
        if not os.path.exists(csv_file):
            logging.warning(f"{csv_file} does not exist in the current directory.")
            print(f"Warning: {csv_file} is missing in the directory.")
            continue

        try:
            print(f"\nLoading {csv_file}...")
            df = pd.read_csv(csv_file)

            # Check if 'NetBIOS' column exists and clean/standardize it
            if 'NetBIOS' in df.columns:
                df['NetBIOS'] = df['NetBIOS'].astype(str).str.strip().str.lower()
            else:
                logging.warning(f"{csv_file} is missing the 'NetBIOS' column.")
                print(f"Warning: {csv_file} does not contain 'NetBIOS' column. Skipping this file.")
                continue

            # Map 'AD' based on 'NetBIOS' column with "Not Found" as default for unmatched rows
            print(f"Matching and adding 'AD' column to {csv_file}...")
            df['AD'] = df['NetBIOS'].map(ad_lookup).fillna("Not Found")

            # Save the updated CSV
            df.to_csv(csv_file, index=False)
            logging.info(f"Successfully updated {csv_file} with AD values.")
            print(f"Successfully added 'AD' column to {csv_file}.")

        except Exception as e:
            logging.error(f"Failed to process {csv_file}: {e}")
            print(f"Error: Could not process {csv_file}. Check log for details.")

    print("AD data processing completed. Check the log file for detailed information.")
    logging.info("Completed adding 'AD' column to all CSV files.")

def add_intune_columns_to_csv_files():
    print("Starting the process to add Intune columns to CSV files based on Intune.csv data...")

    # Load Intune.csv data
    try:
        print("Loading Intune.csv...")
        intune_data = pd.read_csv('Intune.csv', usecols=['Device name', 'Last check-in', 'Serial number', 'Primary user email address'])
        
        # Clean and standardize 'Device name' for consistent matching
        intune_data['Device name'] = intune_data['Device name'].astype(str).str.strip().str.lower()
        
        # Remove duplicates in 'Device name' and keep the first occurrence
        intune_data = intune_data.drop_duplicates(subset='Device name', keep='first')
        
        # Rename 'Last check-in' column for clarity
        intune_data.rename(columns={'Last check-in': 'Last check-in Intune'}, inplace=True)
        
        # Create a lookup dictionary with cleaned 'Device name'
        intune_lookup = intune_data.set_index('Device name').to_dict('index')
        logging.info("Successfully loaded Intune.csv and prepared lookup dictionary.")
    except Exception as e:
        logging.error(f"Failed to load Intune.csv: {e}")
        print("Error: Could not load Intune.csv. Check log for details.")
        return

    # List of CSV files to process
    csv_files = [
        "QDS-above-70-crossed-40d.csv",
        "QDS-0-69-crossed-40d.csv",
        "QDS-0-69-less-40d.csv",
        "QDS-above-70-less-40d.csv"
    ]

    # Process each CSV file with tqdm for a progress bar
    for csv_file in tqdm(csv_files, desc="Processing CSV files for Intune data", unit="file"):
        if not os.path.exists(csv_file):
            logging.warning(f"{csv_file} does not exist in the current directory.")
            print(f"Warning: {csv_file} is missing in the directory.")
            continue

        try:
            print(f"\nLoading {csv_file}...")
            df = pd.read_csv(csv_file)

            # Check if 'NetBIOS' column exists and clean/standardize it
            if 'NetBIOS' in df.columns:
                df['NetBIOS'] = df['NetBIOS'].astype(str).str.strip().str.lower()
            else:
                logging.warning(f"{csv_file} is missing the 'NetBIOS' column.")
                print(f"Warning: {csv_file} does not contain 'NetBIOS' column. Skipping this file.")
                continue

            # Add Intune columns by mapping data based on 'NetBIOS' values, with "Not Found" as default for unmatched rows
            print(f"Matching and adding Intune columns to {csv_file}...")
            df['Last check-in Intune'] = df['NetBIOS'].map(lambda x: intune_lookup.get(str(x).strip().lower(), {}).get('Last check-in Intune', "Not Found"))
            df['Serial number'] = df['NetBIOS'].map(lambda x: intune_lookup.get(str(x).strip().lower(), {}).get('Serial number', "Not Found"))
            df['Primary user email address'] = df['NetBIOS'].map(lambda x: intune_lookup.get(str(x).strip().lower(), {}).get('Primary user email address', "Not Found"))

            # Save the updated CSV
            df.to_csv(csv_file, index=False)
            logging.info(f"Successfully updated {csv_file} with Intune columns.")
            print(f"Successfully added Intune columns to {csv_file}.")

        except Exception as e:
            logging.error(f"Failed to process {csv_file} with Intune data: {e}")
            print(f"Error: Could not process {csv_file} for Intune data. Check log for details.")

    print("Intune data processing completed. Check the log file for detailed information.")
    logging.info("Completed adding Intune columns to all CSV files.")

def add_arch_and_based_columns():
    print("Starting the process to add 'ARCH' and 'BASED' columns to CSV files based on 'Results' column...")

    # List of CSV files to process
    csv_files = [
        "QDS-above-70-crossed-40d.csv",
        "QDS-0-69-crossed-40d.csv",
        "QDS-0-69-less-40d.csv",
        "QDS-above-70-less-40d.csv"
    ]

    # Process each CSV file with tqdm for a progress bar
    for csv_file in tqdm(csv_files, desc="Processing CSV files for ARCH and BASED data", unit="file"):
        if not os.path.exists(csv_file):
            logging.warning(f"{csv_file} does not exist in the current directory.")
            print(f"Warning: {csv_file} is missing in the directory.")
            continue

        try:
            print(f"\nLoading {csv_file}...")
            df = pd.read_csv(csv_file)

            # Check if 'Results' column exists
            if 'Results' not in df.columns:
                logging.warning(f"{csv_file} is missing the 'Results' column.")
                print(f"Warning: {csv_file} does not contain 'Results' column. Skipping this file.")
                continue

            # Convert 'Results' to lowercase to ensure case-insensitive search
            df['Results'] = df['Results'].astype(str).str.lower()

            # Define ARCH based on specific strings in 'Results' column, with "BOTH" for mixed indicators
            df['ARCH'] = df['Results'].apply(
                lambda x: "BOTH" if ("%programfiles%" in x or "c:\\program files\\" in x) and
                                    ("%programfiles(x86)%" in x or "c:\\program files (x86)\\" in x) else
                          "64 Bit" if "%programfiles%" in x or "c:\\program files\\" in x else
                          "32 Bit" if "%programfiles(x86)%" in x or "c:\\program files (x86)\\" in x else
                          "To Check"
            )

            # Define BASED with additional criteria for machine and user profile locations
            df['BASED'] = df['Results'].apply(
                lambda x: "Both locations" if any(loc in x for loc in ["programfiles", "programfiles(x86)", "c:\\program files\\", "c:\\program files (x86)\\"]) and any(user in x for user in ["c:\\users\\", "%systemdrive%\\users\\"]) else
                          "Machine location" if any(loc in x for loc in ["programfiles", "programfiles(x86)", "c:\\program files\\", "c:\\program files (x86)\\"]) else
                          "User profile location" if any(user in x for user in ["c:\\users\\", "%systemdrive%\\users\\"]) else
                          "To Check"
            )

            # Save the updated CSV
            df.to_csv(csv_file, index=False)
            logging.info(f"Successfully updated {csv_file} with ARCH and BASED values.")
            print(f"Successfully added 'ARCH' and 'BASED' columns to {csv_file}.")

        except Exception as e:
            logging.error(f"Failed to process {csv_file} for ARCH and BASED data: {e}")
            print(f"Error: Could not process {csv_file} for ARCH and BASED data. Check log for details.")

    print("ARCH and BASED data processing completed. Check the log file for detailed information.")
    logging.info("Completed adding ARCH and BASED columns to all CSV files.")

if __name__ == '__main__':
    main()
