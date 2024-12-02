# delete_rows_with_empty_cells

This script cleans Excel files (.xlsx) by deleting rows with empty cells in specified columns or in any column, based on user input.

## Overview

The `delete_rows_with_empty_cells` script processes all Excel (.xlsx) files within a user-selected directory. It prompts the user to specify one or more columns (or "all") and then deletes all rows in each file where the specified columns contain empty cells. The new clean files are saved with a "_clean" suffix in the same directory.

## Requirements

- Python 3
- `pandas` library (for data manipulation and exporting to .xlsx)
- `tkinter` library (for file and input dialogs)
- `os` library (for file path operations)

## Files

- `delete_rows_with_empty_cells.py`

## Usage

1. Run the script.
2. A file dialog will prompt you to select a directory containing the Excel files.
3. After selecting the directory, the script will ask you to enter the column letters (e.g., B, C, AC, DE) or column names separated by commas. Alternatively, you can type "all" to delete rows with empty cells in any column.
4. The script will process each Excel file in the selected directory, delete rows based on the specified criteria, and save a new version of the file with a "_clean" suffix in the same directory.

## Important Notes

- Ensure the selected files are valid Excel files in .xlsx format.
- When specifying columns, you can enter column letters (e.g., B, C) or column names. If "all" is entered, rows with any empty cells in any column will be deleted.
- If a specified column is not found or if an invalid column letter is provided, the script will skip processing that column for the file.

## License

This project is governed by the CC BY-NC 4.0 license. For comprehensive details, kindly refer to the LICENSE file included with this project.
