import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog

# Function to prompt user for directory selection
def select_directory():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    folder_selected = filedialog.askdirectory(title="Select the folder containing Excel files")
    return folder_selected

# Function to prompt user for column input
def input_column():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    column_selected = simpledialog.askstring("Input", "Enter the column letter (e.g., B) or column name with empty cells to delete rows:")
    return column_selected

# Main function
def clean_excel_files():
    # Prompt user to select the directory
    directory = select_directory()
    if not directory:
        print("No directory selected.")
        return
    
    # Prompt user to enter the column letter or name
    column_input = input_column()
    if not column_input:
        print("No column selected.")
        return
    
    # Convert column letter to index if necessary
    column_to_check = column_input.strip()

    # List all Excel files in the selected directory
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    
    for file in files:
        # Load the Excel file
        file_path = os.path.join(directory, file)
        df = pd.read_excel(file_path)
        
        # Determine if column_to_check is a letter (convert to column index) or column name
        if len(column_to_check) == 1 and column_to_check.isalpha():  # Assuming column letter is provided (e.g., B)
            column_index = ord(column_to_check.upper()) - ord('A')  # Convert letter to 0-based index
            if column_index >= len(df.columns):
                print(f"Invalid column letter '{column_to_check}' for file {file}. Skipping.")
                continue
            column_to_check = df.columns[column_index]
        elif column_to_check not in df.columns:
            print(f"Column '{column_to_check}' not found in file {file}. Skipping.")
            continue

        # Drop rows where the selected column has empty cells
        df_cleaned = df.dropna(subset=[column_to_check])

        # Generate new file name with "_clean" suffix
        new_file_path = os.path.join(directory, file.replace('.xlsx', '_clean.xlsx'))
        
        # Save the cleaned DataFrame to the new file
        df_cleaned.to_excel(new_file_path, index=False)
        
        print(f"Processed {file} and saved as {new_file_path}")

# Run the main function
if __name__ == "__main__":
    clean_excel_files()
