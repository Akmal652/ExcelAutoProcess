# ExcelAutoProcess
Delivering a client a python-based script automation for pre-processing Excel files in (.xls) or (.xlsx) format and save them in (.csv) format.

Project Description:

This script is designed to process Excel files efficiently by performing various data transformations. It provides a streamlined solution for tasks such as deleting unwanted columns and rows, removing duplicate values, replacing gender identifiers with appropriate first name pronouns, adding country codes to mobile numbers, and handling multiple Excel files in an iterative manner.

Key Features:
- Automatic processing of Excel files: The script can iterate over a folder containing Excel files and process them one by one.
- Column and row manipulation: Unwanted columns and rows can be easily deleted from the Excel files.
- Duplicate value removal: Duplicate values in specific columns can be identified and removed.
- Gender pronoun replacement: Gender identifiers (Male and Female) can be replaced with appropriate first name pronouns (Mr and Miss).
- Mobile number formatting: The script can add country codes to mobile numbers for consistency.
- Error handling: The script detects and handles exceptions, providing detailed error logs for troubleshooting.

This script provides a reliable and efficient solution for automating the processing of Excel files, making data transformations faster and more consistent. It simplifies data cleaning tasks and ensures data integrity for further analysis or usage.


📌 NEW! Added features for excel files processing in v3 :
- Malaysian names recognization & preprocessing: Recognize malaysian names and process any rows that deviate from malaysian names and remove any rows with weird values/symbols in names.


📌 NEW! Script features in v3 :
1. Error logging : Now any errors with corresponding files name producing the errors will be saved in an error log file inside its own directory 'error_logs' with separation between errors.
2. Setup initialization : Script will create 3 folders upon first-time run of script as part of setup initialization process.
3. Animated progress bar : Utilized 'alive-progress' bar by rsalmei to monitor progresss of excel files processing.
