import pandas as pd
import os
import glob
import datetime
import logging
from tqdm import tqdm
import csv  # Import csv module to specify quoting

class TqdmLoggingHandler(logging.Handler):
    """
    Custom logging handler that uses tqdm's write method to prevent interference with the progress bar.
    """
    def emit(self, record):
        try:
            msg = self.format(record)
            tqdm.write(msg)
        except Exception:
            pass

def setup_logging():
    """
    Sets up logging to both a file and the console, retaining only the last 30 days of log entries.
    """
    log_file = 'processing.log'
    
    # Retain only entries from the last 30 days
    if os.path.exists(log_file):
        with open(log_file, 'r') as f:
            lines = f.readlines()
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=30)
        new_lines = []
        for line in lines:
            try:
                timestamp_str = line.split(' - ')[0]
                log_time = datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                if log_time >= cutoff_date:
                    new_lines.append(line)
            except ValueError:
                # If timestamp parsing fails, keep the line
                new_lines.append(line)
        with open(log_file, 'w') as f:
            f.writelines(new_lines)
    
    # Configure the logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Remove existing handlers to prevent duplication
    if logger.hasHandlers():
        logger.handlers.clear()
    
    # File handler
    fh = logging.FileHandler(log_file)
    fh.setLevel(logging.INFO)
    
    # Console handler using TqdmLoggingHandler
    ch = TqdmLoggingHandler()
    ch.setLevel(logging.INFO)
    
    # Formatter with timestamp
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                  datefmt='%Y-%m-%d %H:%M:%S')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    
    # Add handlers to the logger
    logger.addHandler(fh)
    logger.addHandler(ch)

def determine_new_filename(original_filename):
    """
    Determines the new filename based on predefined patterns.
    If no pattern matches, returns None to skip processing.
    """
    patterns = {
        'QDS-above-70-crossed-40d': 'QDS-above-70-crossed-40d.csv',
        'QDS-0-69-crossed-40d': 'QDS-0-69-crossed-40d.csv',
        'QDS-0-69-less-40d': 'QDS-0-69-less-40d.csv',
        'QDS-above-70-less-40d': 'QDS-above-70-less-40d.csv',
    }
    for pattern, new_name in patterns.items():
        if pattern in original_filename:
            return new_name
    # If the file doesn't match any pattern, return None
    return None

def process_csv_files():
    """
    Processes matching CSV files according to specified rules.
    """
    # Record the start time
    start_time = datetime.datetime.now()
    logging.info("Script started at {}".format(start_time.strftime('%Y-%m-%d %H:%M:%S')))
    
    # Load the lookup Excel file
    try:
        logging.info("Loading lookup file: Application Release_Date.xlsx")
        lookup_df = pd.read_excel('Application Release_Date.xlsx', dtype=str)
        required_columns = {'QID', 'Application', 'Release_Date'}
        if not required_columns.issubset(lookup_df.columns):
            logging.error("Lookup file must contain 'QID', 'Application', and 'Release_Date' columns.")
            return
        logging.info("Lookup file loaded successfully.")
    except Exception as e:
        logging.error("Failed to load lookup file: {}".format(e))
        return

    # Define the patterns to look for in filenames
    filename_patterns = [
        'QDS-above-70-crossed-40d',
        'QDS-0-69-crossed-40d',
        'QDS-0-69-less-40d',
        'QDS-above-70-less-40d'
    ]

    # Collect only CSV files that contain the specified patterns
    csv_files = []
    for pattern in filename_patterns:
        matched_files = glob.glob('*{}*.csv'.format(pattern))
        csv_files.extend(matched_files)
        if matched_files:
            logging.info("Found {} matching files for pattern '{}': {}".format(
                len(matched_files), pattern, matched_files))
    
    # Remove duplicates (in case a file matches multiple patterns)
    csv_files = list(set(csv_files))

    if not csv_files:
        logging.info("No matching CSV files found in the current directory.")
        return

    logging.info("Total files to process: {}".format(len(csv_files)))

    # Process each CSV file with a progress bar
    for filepath in tqdm(csv_files, desc='Processing CSV files', unit='file'):
        try:
            logging.info("-----")
            logging.info("Starting processing of file: {}".format(filepath))
            
            # Step 1: Load CSV and remove the first four rows
            logging.info("Loading CSV file and removing the first four rows.")
            df = pd.read_csv(filepath, header=None, dtype=str, skiprows=4, na_filter=False)
            df.reset_index(drop=True, inplace=True)
            logging.info("First four rows removed.")

            # Step 2: Set the first row as header and remove it from data
            logging.info("Setting the first row as header.")
            df.columns = df.iloc[0].str.strip()
            df = df[1:].reset_index(drop=True)
            logging.info("Header set and first data row removed.")

            # Step 3: Drop specific columns by index, including 'Type' column (index 10)
            logging.info("Identifying columns to drop by indices.")
            total_columns = len(df.columns)
            indices_to_drop = [10] + list(range(12, 16)) + list(range(22, 30)) + \
                              list(range(32, 38)) + list(range(39, 43))
            indices_to_drop = [i for i in indices_to_drop if i < total_columns]
            cols_to_drop = [df.columns[i] for i in indices_to_drop]
            logging.info("Columns to drop (by index): {}".format(indices_to_drop))
            logging.info("Columns to drop (by name): {}".format(cols_to_drop))
            df.drop(columns=cols_to_drop, inplace=True, errors='ignore')
            logging.info("Specified columns dropped.")

            # Step 4: Add 'Date' column at the beginning
            current_date = datetime.datetime.now().strftime('%d-%m-%Y')
            logging.info("Adding 'Date' column with today's date: {}".format(current_date))
            df.insert(0, 'Date', current_date)
            df['Date'] = df['Date'].astype(str)  # Ensure 'Date' column is treated as string
            logging.info("'Date' column added.")

            # Step 5: Clean 'OS' column
            if 'OS' in df.columns:
                logging.info("Cleaning 'OS' column.")
                df['OS'] = df['OS'].astype(str).str.strip()
                # Remove rows with specified keywords or blank 'OS'
                keywords = ['ubuntu', 'server', 'linux', 'datacenter']
                pattern = '|'.join(keywords)
                initial_row_count = len(df)
                df = df[~df['OS'].str.contains(pattern, case=False, na=False)]
                df = df[df['OS'] != '']
                final_row_count = len(df)
                logging.info("Filtered 'OS' column. Rows before: {}, after: {}".format(
                    initial_row_count, final_row_count))
            else:
                logging.warning("'OS' column not found in {}.".format(filepath))

            # Step 6: Perform lookup based on 'QID' column
            if 'QID' in df.columns:
                logging.info("Performing lookup based on 'QID' column.")
                df['QID'] = df['QID'].astype(str)
                df = df.merge(lookup_df[['QID', 'Application', 'Release_Date']], on='QID', how='left')
                logging.info("Lookup completed. Filling missing 'Application' and 'Release_Date' with 'Not Found'.")
                df['Application'] = df['Application'].fillna('Not Found')
                df['Release_Date'] = df['Release_Date'].fillna('Not Found')
                logging.info("Missing values filled.")
            else:
                logging.warning("'QID' column not found in {}. Adding 'Application' and 'Release_Date' columns with 'Not Found'.".format(filepath))
                df['Application'] = 'Not Found'
                df['Release_Date'] = 'Not Found'

            # Step 7: Determine new filename and save the processed CSV
            new_filename = determine_new_filename(filepath)
            if new_filename:
                logging.info("Saving processed file as: {}".format(new_filename))
                # Save the DataFrame to CSV with quoting to preserve date format
                df.to_csv(new_filename, index=False, quoting=csv.QUOTE_ALL)
                logging.info("File saved successfully.")
                
                # Step 8: Delete the original file after successful processing
                logging.info("Deleting original file: {}".format(filepath))
                os.remove(filepath)
                logging.info("Original file deleted.")
            else:
                # This should not happen as we filtered files, but just in case
                logging.warning("Filename pattern not matched for {}. Skipping file.".format(filepath))
                continue

        except Exception as e:
            logging.error("Error processing {}: {}".format(filepath, e))
            continue

    # Calculate the total execution time
    end_time = datetime.datetime.now()
    total_time = end_time - start_time
    total_seconds = total_time.total_seconds()
    minutes, seconds = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)
    time_taken = ""
    if hours > 0:
        time_taken += "{} hours ".format(int(hours))
    if minutes > 0:
        time_taken += "{} minutes ".format(int(minutes))
    time_taken += "{:.2f} seconds".format(seconds)

    # Display and log the total execution time
    message = "Total execution time: {}".format(time_taken)
    logging.info(message)
    tqdm.write(message)
    logging.info("Script ended at {}".format(end_time.strftime('%Y-%m-%d %H:%M:%S')))

if __name__ == "__main__":
    setup_logging()
    process_csv_files()








































import pandas as pd
import os
import time
import logging
from tqdm import tqdm
from datetime import datetime, timedelta

# Set up logging configuration
log_file = 'reposition.log'
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to remove old log entries (keeping logs for 30 days)
def clean_log():
    with open(log_file, 'r') as file:
        lines = file.readlines()
    cutoff_date = datetime.now() - timedelta(days=30)
    with open(log_file, 'w') as file:
        for line in lines:
            # Check if line starts with a date
            try:
                log_date = datetime.strptime(line.split(" - ")[0], "%Y-%m-%d %H:%M:%S")
                if log_date >= cutoff_date:
                    file.write(line)
            except ValueError:
                # Write any non-date lines directly
                file.write(line)

# Function to move and update column with timing and error logging
def move_and_update_column(file_path, column_position=11, new_position=21):
    try:
        # Start timing
        start_time = time.time()
        
        # Read the CSV file
        df = pd.read_csv(file_path)
        
        # Ensure the column index is within the DataFrame range
        if column_position >= len(df.columns):
            error_message = f"File '{file_path}' does not have a column at index {column_position}. Skipping."
            print(error_message)
            logging.error(error_message)
            return
        
        # Move the column to the new position
        column_name = df.columns[column_position]
        col_data = df.pop(column_name)
        df.insert(new_position, column_name, col_data)
        
        # Clear all values in the moved column (now at index 21) and rename it to "Type"
        df[column_name] = ""
        df.rename(columns={column_name: "Type"}, inplace=True)
        
        # Update "Type" column based on "OS" values
        if 'OS' in df.columns:
            df['Type'] = df['OS'].apply(lambda x: 'Windows' if pd.notna(x) and 'windows' in x.lower() else
                                                  'Mac' if pd.notna(x) and 'mac' in x.lower() else
                                                  'Not Found')
        else:
            warning_message = f"File '{file_path}' does not contain an 'OS' column. Setting 'Type' to 'Not Found'."
            print(warning_message)
            logging.warning(warning_message)
            df['Type'] = 'Not Found'
        
        # Save the modified DataFrame back to the original file
        df.to_csv(file_path, index=False)
        
        # Log completion and time taken
        end_time = time.time()
        duration = end_time - start_time
        success_message = f"Processed '{file_path}' in {duration:.2f} seconds."
        print(success_message)
        logging.info(success_message)
    except Exception as e:
        # Log any exceptions that occur
        error_message = f"An error occurred while processing '{file_path}': {str(e)}"
        print(error_message)
        logging.error(error_message)

# Function to rename column header "Release_Date" to "Publish Date"
def rename_column(file_path, old_column_name="Release_Date", new_column_name="Publish Date"):
    try:
        # Read the CSV file
        df = pd.read_csv(file_path)
        
        # Check if the old column exists
        if old_column_name in df.columns:
            # Rename the column
            df.rename(columns={old_column_name: new_column_name}, inplace=True)
            # Save changes to the file
            df.to_csv(file_path, index=False)
            print(f"Renamed '{old_column_name}' to '{new_column_name}' in '{file_path}'.")
            logging.info(f"Renamed '{old_column_name}' to '{new_column_name}' in '{file_path}'.")
        else:
            warning_message = f"File '{file_path}' does not contain the column '{old_column_name}'."
            print(warning_message)
            logging.warning(warning_message)
    except Exception as e:
        error_message = f"An error occurred while renaming column in '{file_path}': {str(e)}"
        print(error_message)
        logging.error(error_message)

# Define the file names to process
file_names = [
    'QDS-above-70-crossed-40d.csv',
    'QDS-0-69-crossed-40d.csv',
    'QDS-0-69-less-40d.csv',
    'QDS-above-70-less-40d.csv'
]

# Clean old log entries
clean_log()

# Process each file with progress display
for file_name in tqdm(file_names, desc="Processing Files", unit="file"):
    # Move and update the Type column
    move_and_update_column(file_name)
    # Rename Release_Date to Publish Date
    rename_column(file_name)

# Print completion message
print("All files processed. Check 'reposition.log' for detailed logs.")
































































































import os
import pandas as pd

def update_csv_dates():
    # List all CSV files in the current working directory
    csv_files = [file for file in os.listdir() if file.endswith('.csv')]
    
    # Prompt user to enter the new date
    new_date = input("Enter the new date (dd-mm-yyyy): ")
    
    # Validate date format
    if not pd.to_datetime(new_date, format='%d-%m-%Y', errors='coerce'):
        print("Invalid date format. Please use dd-mm-yyyy.")
        return

    # Process each CSV file
    for file in csv_files:
        df = pd.read_csv(file)
        
        # Check if the first column is named 'Date'
        if df.columns[0] == 'Date':
            # Replace all values in the 'Date' column with the new date
            df['Date'] = new_date
            
            # Save the modified DataFrame back to the CSV file, retaining the header
            df.to_csv(file, index=False)
            print(f"Updated {file} with the new date.")
        else:
            print(f"Skipped {file} as it does not have 'Date' as the first column.")
            
    print("All files processed.")

# Run the function
update_csv_dates()





































import pandas as pd
from datetime import datetime, timedelta

# Load your data from Power BI into a DataFrame
df = YourTableName  # Power BI will automatically replace this with the table data

# Convert the Date column to datetime if it isn't already
df['Date'] = pd.to_datetime(df['Date'])

# Filter for dates within the last 5 days, including today
today = datetime.today().date()
five_days_ago = today - timedelta(days=4)  # Adjusted to get today and the past 4 days
filtered_df = df[(df['Date'] >= pd.Timestamp(five_days_ago)) & (df['Date'] <= pd.Timestamp(today))]

# Drop the 'Total' column if it exists
if 'Total' in filtered_df.columns:
    filtered_df = filtered_df.drop(columns=['Total'])

# Output the filtered DataFrame back to Power BI
dataset = filtered_df












































































































import pandas as pd
from tqdm import tqdm
import logging
from logging.handlers import TimedRotatingFileHandler
import time
import os

# List of CSV files to process
file_names = [
    'QDS-above-70-crossed-40d.csv',
    'QDS-0-69-crossed-40d.csv',
    'QDS-0-69-less-40d.csv',
    'QDS-above-70-less-40d.csv'
]

def setup_logger():
    """Sets up the logger to log to 'date.log' and rotate logs after 30 days."""
    logger = logging.getLogger('DateConversionLogger')
    logger.setLevel(logging.INFO)
    
    # Create a handler that rotates logs every day and keeps 30 backups
    handler = TimedRotatingFileHandler('date.log', when='D', interval=1, backupCount=30)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    
    # Also log to console
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    logger.addHandler(console)
    
    return logger

def convert_dates_in_file(file_name, logger):
    """Converts the 'Published Date' column to 'dd-mm-yyyy' format."""
    logger.info(f"Processing file: {file_name}")
    start_time = time.time()
    
    # Read the CSV file
    df = pd.read_csv(file_name)
    logger.info("CSV file read into DataFrame.")
    
    # Check if 'Published Date' column exists
    if 'Published Date' not in df.columns:
        logger.error(f"'Published Date' column not found in {file_name}. Skipping file.")
        return
    
    # Convert the 'Published Date' column
    logger.info("Converting 'Published Date' column to 'dd-mm-yyyy' format.")
    df['Published Date'] = pd.to_datetime(df['Published Date'], errors='coerce')
    df['Published Date'] = df['Published Date'].dt.strftime('%d-%m-%Y')
    
    # Save the DataFrame back to CSV
    df.to_csv(file_name, index=False)
    logger.info(f"File saved after date conversion: {file_name}")
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    logger.info(f"Time taken to process {file_name}: {elapsed_time:.2f} seconds.\n")

def main():
    logger = setup_logger()
    total_start_time = time.time()
    logger.info("Script started.")
    
    for file_name in tqdm(file_names, desc="Processing files"):
        convert_dates_in_file(file_name, logger)
    
    total_end_time = time.time()
    total_elapsed_time = total_end_time - total_start_time
    logger.info(f"Total time taken for script: {total_elapsed_time:.2f} seconds.")
    logger.info("Script finished.")

if __name__ == "__main__":
    main()

























import os
import time
import logging
from tqdm import tqdm
from datetime import datetime, timedelta

def clean_old_logs(log_filename, days=30):
    """Keep log entries from the last 'days' days."""
    if not os.path.exists(log_filename):
        return
    with open(log_filename, 'r') as f:
        lines = f.readlines()
    cutoff_date = datetime.now() - timedelta(days=days)
    new_lines = []
    for line in lines:
        date_str = line.split(' - ')[0]
        try:
            log_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S,%f')
            if log_date >= cutoff_date:
                new_lines.append(line)
        except ValueError:
            # Keep the line if date parsing fails
            new_lines.append(line)
    with open(log_filename, 'w') as f:
        f.writelines(new_lines)

def setup_logging(log_filename):
    """Set up logging to console and log file."""
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # File handler
    fh = logging.FileHandler(log_filename, mode='a')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

def main():
    file_names = [
        'QDS-above-70-crossed-40d.csv',
        'QDS-0-69-crossed-40d.csv',
        'QDS-0-69-less-40d.csv',
        'QDS-above-70-less-40d.csv'
    ]

    log_filename = 'clean.log'
    clean_old_logs(log_filename, days=30)
    setup_logging(log_filename)
    logger = logging.getLogger()

    total_start_time = time.time()
    logger.info('Script started.')

    for file in tqdm(file_names, desc='Processing files'):
        step_start_time = time.time()
        if os.path.exists(file):
            try:
                os.remove(file)
                logger.info(f"Deleted file: {file}")
            except Exception as e:
                logger.error(f"Error deleting file {file}: {e}")
        else:
            logger.info(f"File not found: {file}, nothing to delete.")
        step_end_time = time.time()
        elapsed_time = step_end_time - step_start_time
        logger.info(f"Time taken for {file}: {elapsed_time:.2f} seconds.")

    total_end_time = time.time()
    total_elapsed_time = total_end_time - total_start_time
    logger.info(f"Script completed in {total_elapsed_time:.2f} seconds.")

if __name__ == "__main__":
    main()





















import subprocess
import colorlog
import logging
import sys
import time
import threading
import queue

# Step 1: Setup Color Logger
def setup_logger():
    """Sets up a logger to use colorlog for colored console output."""
    logger = logging.getLogger('WrapperLogger')
    logger.setLevel(logging.INFO)

    # Create a console handler using colorlog
    console_handler = colorlog.StreamHandler()
    console_formatter = colorlog.ColoredFormatter(
        '%(log_color)s%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        reset=True,  # Ensure the color is reset after each log message
        log_colors={
            'DEBUG': 'cyan',
            'INFO': 'green',
            'WARNING': 'yellow',
            'ERROR': 'red',
            'CRITICAL': 'bold_red',
        }
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    return logger

# Step 2: Function to process lines from a stream (stdout or stderr)
def enqueue_output(pipe, queue):
    for line in iter(pipe.readline, ''):
        queue.put(line)
    pipe.close()

# Step 3: Run Scripts in Sequence with Keyword-Based Coloring
def run_scripts(scripts, logger):
    start_time = time.time()  # Track the overall start time
    
    # Color mapping for keywords
    color_mapping = {
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'bold_red'
    }

    for script in scripts:
        logger.info(f"Running {script}...")
        script_start_time = time.time()  # Track individual script start time
        try:
            # Run the script and capture stdout and stderr
            process = subprocess.Popen(
                [sys.executable, script],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1  # Line-buffered for real-time output processing
            )

            # Create queues to handle stdout and stderr lines
            q_stdout = queue.Queue()
            q_stderr = queue.Queue()

            # Start threads to read from stdout and stderr
            stdout_thread = threading.Thread(target=enqueue_output, args=(process.stdout, q_stdout))
            stderr_thread = threading.Thread(target=enqueue_output, args=(process.stderr, q_stderr))
            stdout_thread.daemon = True
            stderr_thread.daemon = True
            stdout_thread.start()
            stderr_thread.start()

            # Process the output as it becomes available
            while True:
                # Handle stdout
                try:
                    line = q_stdout.get_nowait()
                except queue.Empty:
                    line = None
                if line:
                    colorize_and_log(line.strip(), logger, color_mapping)

                # Handle stderr
                try:
                    line = q_stderr.get_nowait()
                except queue.Empty:
                    line = None
                if line:
                    colorize_and_log(line.strip(), logger, color_mapping, is_error=True)

                # Exit loop if process is done and queues are empty
                if process.poll() is not None and q_stdout.empty() and q_stderr.empty():
                    break

            # Check return code after the script ends
            if process.returncode == 0:
                elapsed_time = time.time() - script_start_time
                logger.info(f"{script} finished successfully in {format_time(elapsed_time)}.")
            else:
                logger.error(f"{script} failed with return code {process.returncode}.")

        except Exception as e:
            logger.error(f"An exception occurred while running {script}: {str(e)}")

    # Calculate and log total time taken for all scripts
    total_elapsed_time = time.time() - start_time
    logger.info(f"Total time taken for all scripts: {format_time(total_elapsed_time)}.")

# Step 4: Helper function to colorize output based on keywords
def colorize_and_log(line, logger, color_mapping, is_error=False):
    """Logs a line with appropriate color based on keywords."""
    for keyword, color in color_mapping.items():
        if keyword in line:
            # Log with appropriate level and color
            log_color = color_mapping.get(keyword, 'green')
            log_message = f'{line}'
            colored_formatter = colorlog.ColoredFormatter(
                f'%(log_color)s%(message)s',
                reset=True,
                log_colors={keyword: log_color}
            )
            console_handler = colorlog.StreamHandler()
            console_handler.setFormatter(colored_formatter)
            temp_logger = logging.getLogger('TempLogger')
            temp_logger.addHandler(console_handler)
            temp_logger.setLevel(logging.INFO)

            if keyword in ['ERROR', 'CRITICAL']:
                temp_logger.error(log_message)
            else:
                temp_logger.info(log_message)

            temp_logger.removeHandler(console_handler)
            return

    # If no keywords match, log as info by default
    if is_error:
        logger.error(line)
    else:
        logger.info(line)

# Step 5: Define the list of scripts to be run
scripts_to_run = [
    "clean.py",
    "del.py",
    "reposition.py",
    "validate.py",
    "PublishDateFormatt.py"
]

# Step 6: Helper function to format time
def format_time(seconds):
    """Formats time in seconds to a readable format (e.g., 5 min 39 sec)."""
    minutes = int(seconds // 60)
    seconds = int(seconds % 60)
    if minutes > 0:
        return f"{minutes} min {seconds} sec"
    else:
        return f"{seconds} sec"

# Step 7: Run the wrapper
if __name__ == "__main__":
    # Set up the logger
    logger = setup_logger()
    
    # Run each script in sequence
    run_scripts(scripts_to_run, logger)
