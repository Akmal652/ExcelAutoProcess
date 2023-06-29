import gc
import os
import pandas as pd
import sys
import subprocess
import numpy as np
import win32com.client
import datetime
import time
import re
from alive_progress import alive_bar
from loguru import logger
from pywintypes import com_error
from halo import Halo

# Get the current timestamp without seconds
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")

# Get the user's "Documents" folder
documents_folder = os.path.expanduser("~\Documents")

# Input and output folders
input_folder = os.path.join(documents_folder, "Narqes Database", "database_before")
subfolder1 = os.path.basename(os.path.dirname(input_folder))
subfolder2 = os.path.basename(input_folder)
output_folder = os.path.join(documents_folder, "Narqes Database", "database_after")

# Error log folder path
error_log_folder = os.path.join(documents_folder, "Narqes Database", "error_logs")

# Deleting empty error log files
# Check if the error log folder exists
if os.path.exists(error_log_folder):
    # Iterate over the files in the error log folder
    for filename in os.listdir(error_log_folder):
        file_path = os.path.join(error_log_folder, filename)
        # Check if the error log file is empty
        if os.path.isfile(file_path) and os.path.getsize(file_path) == 0:
            try:
                # Delete the empty error log file
                os.remove(file_path)
            except Exception as e:
                print(f"Failed to delete log file '{file_path}': {e}")

# Create input, output, and error logs folders if they don't exist
required_folders = [input_folder, output_folder, error_log_folder]

setup_completed = False

# Create a spinner animation for software setup process
spinner_setup = Halo(text="Starting software initial setup...", spinner="dots")

if all(not os.path.exists(folder) for folder in required_folders):
    # Start the spinner
    spinner_setup.start()
    time.sleep(5)
    with alive_bar(len(required_folders), enrich_print=False, title='Initializing software setup', max_cols = 110, receipt_text = False) as bar:
        for index, folder in enumerate(required_folders):
            os.makedirs(folder)
            bar()
    setup_completed = True

# Verify if the folders exist
folders_exist = all(os.path.exists(f) for f in required_folders)

if setup_completed and folders_exist:
    print("\n")
    # Stop the spinner when setup is complete
    spinner_setup.succeed("Setup has completed. Please place Excel files in the following folders before running the software:")
    print("\n")
    print(f"- {input_folder}\n\n")
    input("Press any key to exit the installation process...")
    sys.exit()

            
# Disable the default console output
logger.remove()

# Configure logging to save exceptions to a file in the error log folder
log_filename = os.path.join(error_log_folder, f"process_database_error_trace_{timestamp}.log")
logger.add(log_filename, level="INFO", colorize=False, format="{message}")
logger.add(log_filename, level="ERROR", colorize=False, format="{time} - {level} - {message}")

# Check if pywin32 package is installed, and install it if necessary
try:
    import win32com.client
except ImportError:
    print("pywin32 package is not found. Installing pywin32...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])

# Check if there are any Excel files in the input folder
excel_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') or f.endswith('.xls')]
if not excel_files:
    print("No Excel files found in the input folder. Exiting the program.")
    sys.exit()

# List to store files that were not processed correctly
unprocessed_files = []

spinner = Halo(text='Starting to process excel files', spinner='dots')
spinner.start()

time.sleep(3)

spinner.stop()

print('\n')
print('Time remaining (eta) can be seen inside progress bar after (~)')
print('\n')

with alive_bar(len(excel_files), enrich_print = False, title='Processing Excel Files' ,max_cols = 120, receipt_text = False) as bar:
    # Iterate over each Excel file
    for excel_file in excel_files:
        # Getting Excel file path
        file_path = os.path.join(input_folder, excel_file)

        excel_app = None
        try:
            # Read the converted file
            if excel_file.endswith('.xls'):
                # Check if pywin32 package is imported successfully
                try:
                    import win32com.client
                except ImportError:
                    print("pywin32 package is not found. Unable to convert .xls to .xlsx.")
                    sys.exit()

                excel_app = win32com.client.Dispatch("Excel.Application")
                try:
                    excel_app.Visible = False
                    excel_app.DisplayAlerts = False  # Disable display alerts
                    if excel_app is not None:
                        excel_app.DisplayAlerts = False  # Double-checking to make sure display alerts are disabled
                
                except AttributeError as e:
                    if "Visible" in str(e):
                        print("Warning: Unable to set the 'Visible' property of Excel.Application.")
                        print('\n')
                        print('Please close Microsoft Excel product activation alert and re-run.')
               

                try:
                    workbook = excel_app.Workbooks.Open(file_path)

                    # Assigning new file name for converted xls files to xlsx
                    new_file_path = os.path.join(input_folder, 'converted_' + os.path.splitext(excel_file)[0] + '.xlsx')

                    workbook.SaveAs(new_file_path, FileFormat=51)  # 51 represents .xlsx format
                    workbook.Close(SaveChanges=False)
              
                except AttributeError as e:
                    print("Warning: Unable to set the 'Visible' property of Excel.Application.")
                    print('\n')
                    print('Please close Microsoft Excel product activation alert and re-run.')

                # Read the converted .xlsx file
                df = pd.read_excel(new_file_path, header=5)

                # Remove the converted file
                os.remove(new_file_path)
            elif excel_file.endswith('.xlsx'):
                # Read the .xlsx file
                df = pd.read_excel(file_path, header=5)
            else:
                print(f"Unsupported file format: {excel_file.suffix}. Skipping file.")
                continue
        except KeyError as e:
            unprocessed_files.append((excel_file, f"KeyError while reading file before processing: {e}."))
            try:
                if excel_app is not None:
                    excel_app.Quit()
            except com_error as e:
                if e.hresult == -2147418111:
                    print(f"Program stopped processing files at {excel_file}")
                    print('\n')
                    print('Please close Microsoft Excel product activation alert and re-run.')
            continue
        except Exception as e:
            unprocessed_files.append((excel_file, f"Error reading file before processing: {str(e)}."))
            try:
                if excel_app is not None:
                    excel_app.Quit()
            except com_error as e:
                if e.hresult == -2147418111:
                    print(f"Program stopped processing files at {excel_file}")
                    print('\n')
                    print('Please close Microsoft Excel product activation alert and re-run.')
            continue

        try:
            # Check and remove rows that contain NaN values in all columns of a DataFrame
            df = df.dropna(how='all')

            # Check rows containing column headers and make the row as column headers and remove the row
            # Check if 'FirstName' column is already present in the column headers
            if 'FirstName' not in df.columns:
                # Find the row that contains the column names
                header_row = df[df.apply(lambda row: 'FirstName' in row.values, axis=1)]

                # Set the column headers and remove the unnecessary row
                df.columns = header_row.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)

            # Task 2 : Delete columns B-D,G,I-X
            desired_columns = ['FirstName', 'Gender', 'Type', 'Mobile']
            existing_columns = df.columns.tolist()

            # Task 2 (a) : Find the common columns
            # Filter the common columns between desired_columns and existing_columns
            common_columns = [col for col in desired_columns if col in existing_columns]

            # Task 2 (b) : Select only the common columns in the DataFrame
            df = df[common_columns]

            if 'Type' in df.columns:
                # Task 3: Delete rows containing the string "Employee"
                df = df[~df['Type'].str.contains('Employee')]
                # Task 3.5: Delete column 'Type'
                df.drop(['Type'], axis=1, inplace=True)

            # Task 4: Replace "Female" with "Ms."
            df['Gender'] = df['Gender'].replace('Female', 'Ms.')

            # Task 5: Replace "Male" with "En."
            df['Gender'] = df['Gender'].replace('Male', 'En.')

            # Task 6: Change phone number format
            df['Mobile'] = df['Mobile'].str.replace('-', '')
            df['Mobile'] = df['Mobile'].str.replace(' ', '')
            df['Mobile'] = '60' + df['Mobile']

            # Task 7: Remove duplicate values and empty values in column "Mobile"
            df.drop_duplicates(subset=['Mobile'], inplace=True)
            df['Mobile'].replace('', np.nan, inplace=True)
            df.dropna(subset=['Mobile'], inplace=True)

            # Function to process the names
            def process_name(name):
                name = re.sub(r'^\.*', '', name)  # Remove leading dots
                if re.match(r'^[a-zA-Z\s\'.]*$', name):
                    return name
                return ''

            # Task 8: Remove rows with names consisting of only special characters
            df['FirstName'] = df['FirstName'].apply(process_name)
            df = df[df['FirstName'] != '']

            # Reset the DataFrame index
            df = df.reset_index(drop=True)

            # Task 8.5: Change values under 'FirstName' column to become proper-cased
            df['FirstName'] = df['FirstName'].str.title()

            # Get the output file path
            output_file = os.path.splitext(excel_file)[0] + '.csv'
            output_path = os.path.join(output_folder, output_file)

            # Save the DataFrame as a CSV file using the original name of the Excel file
            df.to_csv(output_path, index=False)

            print(f"Processed {excel_file} and saved the output to {output_file}")
            print("", end="\n")

        except KeyError as e:
            # Get the current excel file that caused the exception
            current_excel_file = excel_file if 'excel_file' in locals() else ""
            # Get the current timestamp with seconds
            current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")   
            unprocessed_files.append((excel_file, f"KeyError while processing file: Column {e} is missing or deleted during processing."))
            # Log the exception with the corresponding excel file and timestamp
            logger.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
            logger.exception(f"{e}")   # Log the exception inside error trace logs
            logger.error(f"End of error logging for excel file: {current_excel_file}")  # Marks end of error trace log for the corresponding excel file
            logger.info("\n\n") # Log a whole line of space between error trace logs
            logger.info("=" * 170)
            # logging.info('\n')
        except IndexError as e:
            current_excel_file = excel_file if 'excel_file' in locals() else ""
            current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
            unprocessed_files.append((excel_file, f"Index Error while processing file: {e}. {excel_file} excel file is suspected as corrupted file, \nplease re-save {excel_file} excel file in (.xlsx) format."))
            logger.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
            logger.exception(f"{e}")
            logger.error(f"End of error logging for excel file: {current_excel_file}")
            logger.info("\n\n")
            logger.info("=" * 170)
        except Exception as e:
            current_excel_file = excel_file if 'excel_file' in locals() else ""
            current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
            unprocessed_files.append((excel_file, str(e)))
            logger.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
            logger.exception(f"{e}")
            logger.error(f"End of error logging for excel file: {current_excel_file}")
            logger.info("\n\n")
            logger.info("=" * 170)

        gc.collect()

        # Update and increment the progress bar
        bar()  
        
# Display the list of unprocessed files with exceptions
if unprocessed_files:
    print("", end="\n")
    print(f"\n{len(unprocessed_files)} Excel files are unprocessed with their errors listed:")
    for idx, (file, exception) in enumerate(unprocessed_files, 1):
        print("\n")
        print(f"{idx}. {file}: {exception}")
    print("\n\nPlease check the unprocessed files for errors and make sure they are exported from Zenoti correctly, preferably in (.xlsx) or (.xls) format.")
else:
    folder_name = os.path.basename(input_folder)
    print("\n")
    print(f"All excel files inside {folder_name} folder are processed successfully without errors.")

print("\n")

spinner.succeed("Done processing database...")

# Wait for user input before exiting
input("\nPress any key to exit...")
