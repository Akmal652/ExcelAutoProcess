import gc
import os
import pandas as pd
import sys
import subprocess
import numpy as np
import win32com.client
import logging
import datetime

# Get the current timestamp without seconds
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")

# Get the user's "Documents" folder
documents_folder = os.path.expanduser("~\Documents")

# Input and output folders
input_folder = os.path.join(documents_folder, "Narqes Database", "database_before")
output_folder = os.path.join(documents_folder, "Narqes Database", "database_after")

# Error log folder path
error_log_folder = os.path.join(documents_folder, "Narqes Database", "error_logs")

# Create input and output folders if they don't exist
if not os.path.exists(input_folder):
    os.makedirs(input_folder)

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create the error log folder if it doesn't exist
if not os.path.exists(error_log_folder):
    os.makedirs(error_log_folder)

# Configure logging to save exceptions to a file in the error log folder
log_filename = os.path.join(error_log_folder, f"process_database_error_trace_{timestamp}.log")
logging.basicConfig(filename=log_filename, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

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
            excel_app.Visible = False
            excel_app.DisplayAlerts = False  # Disable display alerts
            if excel_app is not None:
                excel_app.DisplayAlerts = False  # Double-checking to make sure display alerts are disabled
            workbook = excel_app.Workbooks.Open(file_path)

            # Assigning new file name for converted xls files to xlsx
            new_file_path = os.path.join(input_folder, 'converted_' + os.path.splitext(excel_file)[0] + '.xlsx')

            workbook.SaveAs(new_file_path, FileFormat=51)  # 51 represents .xlsx format
            workbook.Close(SaveChanges=False)

            # Read the converted .xlsx file
            df = pd.read_excel(new_file_path, header=6)

            # Remove the converted file
            os.remove(new_file_path)
        elif excel_file.endswith('.xlsx'):
            # Read the .xlsx file
            df = pd.read_excel(file_path, header=6)
        else:
            print(f"Unsupported file format: {excel_file.suffix}. Skipping file.")
            continue
    except KeyError as e:
        unprocessed_files.append((excel_file, f"KeyError while reading file before processing: {e}."))
        if excel_app is not None:
            excel_app.Quit()
        continue
    except Exception as e:
        unprocessed_files.append((excel_file, f"Error reading file before processing: {str(e)}."))
        if excel_app is not None:
            excel_app.Quit()
        continue

    try:
        # Rename first row as column headers
        df = df.rename(columns=df.iloc[0]).loc[1:]

        # Task 2: Delete columns B-D, and G
        df.drop(df.columns[[1, 2, 3, 6]], axis=1, inplace=True)

        # Task 2.5: Delete columns I-X
        df.drop(df.columns[4:20], axis=1, inplace=True)

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

        # Get the output file path
        output_file = os.path.splitext(excel_file)[0] + '.csv'
        output_path = os.path.join(output_folder, output_file)

        # Save the DataFrame as a CSV file using the original name of the Excel file
        df.to_csv(output_path, index=False)

        print(f"Processed {excel_file} and saved the output to {output_file}")

    except KeyError as e:
        # Get the current excel file that caused the exception
        current_excel_file = excel_file if 'excel_file' in locals() else ""
        # Get the current timestamp with seconds
        current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")   
        unprocessed_files.append((excel_file, f"KeyError while processing file: Column {e} is missing or deleted during processing."))
        # Log the exception with the corresponding excel file and timestamp
        logging.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
        logging.exception(f"{e}")   # Log the exception inside error trace logs
        logging.error(f"End of error logging for excel file: {current_excel_file}")  # Marks end of error trace log for the corresponding excel file
        logging.error("\n") # Add spacing between error trace log
    except IndexError as e:
        current_excel_file = excel_file if 'excel_file' in locals() else ""
        current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
        unprocessed_files.append((excel_file, f"Index Error while processing file: {e}. {excel_file} excel file is suspected as corrupted file, \nplease re-save {excel_file} excel file in (.xlsx) format."))
        logging.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
        logging.exception(f"{e}")
        logging.error(f"End of error logging for excel file: {current_excel_file}")
        logging.error("\n") 
    except Exception as e:
        current_excel_file = excel_file if 'excel_file' in locals() else ""
        current_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
        unprocessed_files.append((excel_file, str(e)))
        logging.error(f"Exception occurred at {current_timestamp} while processing file: {current_excel_file}")
        logging.exception(f"{e}")
        logging.error(f"End of error logging for excel file: {current_excel_file}")
        logging.error("\n")  

    gc.collect()

# Display the list of unprocessed files with exceptions
if unprocessed_files:
    print(f"\n{len(unprocessed_files)} Excel files are unprocessed with their errors listed:")
    for idx, (file, exception) in enumerate(unprocessed_files, 1):
        print("\n")
        print(f"{idx}. {file}: {exception}")
    print("\n\nPlease check the unprocessed files for errors and make sure they are exported from Zenoti correctly, preferably in (.xlsx) or (.xls) format.")
else:
    folder_name = os.path.basename(input_folder)
    print(f"All excel files inside {folder_name} folder are processed successfully without errors.")

print("\nDone processing database...")

# Wait for user input before exiting
input("\nPress any key to exit...")
