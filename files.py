# Importing libraries and packages
import glob
import os
import shutil
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog
# Library to get rid of the annoying "UserWarning: Workbook contains no default style, apply openpyxl's default"
import warnings
warnings.simplefilter("ignore")
# done

# Let's the user select an Excel file. Downloads folder is the default one
def select_file(instructions = "Select a file", standard_directory = os.path.expanduser('~/Downloads')):
    
    # Create a hidden root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Open the file dialog and ask the user to select a file
    file_path = filedialog.askopenfilename(initialdir = standard_directory, title = instructions)
    
    # Check if the user selected a file
    if file_path:
        print(f"Selected file: {file_path}")
    else:
        print("No file selected.")

    # Load the file into a DataFrame
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(file_path)

    return df

# Gets the latest file from the Downloads folder. CSV or XLSX
def get_file(file_name):

    # Define the start of the filename you’re looking for
    filename_prefix = file_name

    # Path to the Downloads folder
    downloads_path = os.path.expanduser('~/Downloads')

    # Search for files that start with 'filename_' in the downloads folder
    matching_files = glob.glob(os.path.join(downloads_path, f"{filename_prefix}*"))

    # Ensure there's at least one match, and select the most recent if there are multiple
    if matching_files:
        # If there are multiple files, you may want to sort them by modification time
        latest_file = max(matching_files, key=os.path.getmtime)
        print(f"Latest file: {latest_file}")

        # Load the file into a DataFrame
        if latest_file.endswith('.csv'):
            df = pd.read_csv(latest_file)
        elif latest_file.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(latest_file)
        
        return df
    
    else:
        print("No files found matching the specified prefix.")

# Optimizes column headers for sql
def sql_headers(df):
    
    replace_chars = str.maketrans({
            ' ': '_',
            '-': '',
            '?': ''
    })

    df.columns = df.columns.str.lower().str.translate(replace_chars)

    return df

# Saves the Pandas DataFrame in XLSX format
def save_dataframe(
        dataframe, ask_location = False, standard_directory = os.path.expanduser('~/Downloads'), instructions = "Select the desired location to save the dataframe",
        *, format = "xlsx",
    ):
    
    if ask_location == True:
        # Create a hidden root window
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        # Open the file dialog for saving the file
        file_path = filedialog.asksaveasfilename(
            initialdir = standard_directory,
            title = instructions,
            defaultextension = ".xlsx",
            filetypes = [("Excel files", "*.xlsx"), ("CSV documents", "*.csv*"), ("All files", "*.*")]
        )
    
    else:
        file_path = standard_directory
    
    # Check if a file path was provided
    if file_path:
        try:
            # Save the DataFrame to the specified path
            if format == ".xlsx":
                dataframe.to_excel(file_path, index = False)
                print(f"DataFrame successfully saved to {file_path}")

            elif format == ".csv":
                dataframe.to_csv(file_path, index = False)
                print(f"DataFrame successfully saved to {file_path}")

            else:
                print("An error occurred. Your file hasn't been saved.")        
        
        except Exception as e:
            print(f"An error occurred while saving the file: {e}")
        
    else:
        print("No file was selected to save the DataFrame.")

# Create or replaces the folder based on title + today's date
def create_or_replace_folder(title):
    today = datetime.today().strftime('%m-%d-%Y')
    folder_name = f"{title} {today}"
    
    # Check if the folder exists
    if os.path.exists(folder_name):
        # Remove the existing folder
        shutil.rmtree(folder_name)
    
    # Create the new folder
    os.makedirs(folder_name)
    
    return folder_name

# Slices the Excel file based on a specific column
def slice_excel_by_parent(df, folder_name, *, column_to_slice_by, columns_to_remove = None):
    
    try:
        # Check if the slicing column exists
        if column_to_slice_by not in df.columns:
            print(f"Error: The column {column_to_slice_by} does not exist in the Excel file.")
            return False
        
        # Group the data based on the slicing column and save each group as a new Excel file
        for parent_value, group in df.groupby('Parent'):
            # Remove the specified columns from the group
            group = group.drop(columns=[col for col in columns_to_remove if col in group.columns])
            
            # Create the output file name
            output_file = os.path.join(folder_name, f"{parent_value}.xlsx")
            
            # Save the group to an Excel file
            group.to_excel(output_file, index=False)
            
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False
    
    return True