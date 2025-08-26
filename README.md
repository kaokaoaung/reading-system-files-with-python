📂 File List Exporter (Python Script)

This script is designed to scan a folder (including all subfolders) and export a list of all files into an Excel file. It is useful for quickly checking how many files are stored under a specific directory, along with their exact folder paths.

🚀 Features

Recursively scans all files in the target folder.

Records file name and folder path for each file.

Exports results to an Excel file (.xlsx).

Automatically creates the output directory if it does not exist.

Provides a file count after export.

🛠️ Requirements

Make sure you have Python 3 installed, along with the following libraries:

pip install pandas, openpyxl


📊 Example Output

| File Name   | Folder Path                   |
| ----------- | ----------------------------- |
| file_1.xlsx | C:\Users\Me\Documents\Reports |
| file_2.csv    | C:\Users\Me\Documents\Data    |
| image.png   | C:\Users\Me\Documents\Images  |


📌 Use Cases

Quickly auditing how many files exist in a large directory.

Creating an inventory of project files.

Preparing data for file management or migration tasks.
