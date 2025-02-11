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

def read_file(file_path):
    
    # Load the file into a DataFrame
    if file_path.endswith('.csv'):
       df = pd.read_csv(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
       df = pd.read_excel(file_path)
    
    return df

def select_file(instructions = "Select a file", standard_directory = os.path.expanduser('~/Downloads')):
# Let's the user select an Excel file. Downloads folder is the default one

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

    return read_file(file_path)

def get_file(file_name, directory = os.path.expanduser('~/Downloads')):
# Gets the latest file from a folder specified by the user. Downloads folder is the standard directory. csv or xlsx.

    # Define the start of the filename you’re looking for
    filename_prefix = file_name

    # Search for files that start with 'filename_' in the downloads folder
    matching_files = glob.glob(os.path.join(directory, f"{filename_prefix}*"))

    # Ensure there's at least one match, and select the most recent if there are multiple
    if matching_files:
        
        # If there are multiple files, you may want to sort them by modification time
        latest_file = max(matching_files, key=os.path.getmtime)
        print(f"Latest file: {latest_file}")      
        return read_file(latest_file)
    
    else:
        print("No files found matching the specified prefix.")

def sql_headers(df):
# Optimizes column headers for sql
    
    replace_chars = str.maketrans({
            ' ': '_',
            '-': '',
            '?': ''
    })

    df.columns = df.columns.str.lower().str.translate(replace_chars)

    return df

def save_dataframe(
# Saves the Pandas DataFrame in xlsx format

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

def select_folder(instructions = "Select a folder", standard_directory = os.path.expanduser('~/Downloads')):
    
    # Create a hidden root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Open the file dialog and ask the user to select a file
    folder_path = filedialog.askdirectory(initialdir = standard_directory, title = instructions)
    
    # Check if the user selected a file
    if folder_path:
        print(f"Selected folder: {folder_path}")
    else:
        print("No folder selected.")
    
    return folder_path

def create_or_replace_folder(title):
# Create or replaces the folder based on title + today's date

    today = datetime.today().strftime('%m-%d-%Y')
    folder_name = f"{title} {today}"
    
    # Check if the folder exists
    if os.path.exists(folder_name):
        # Remove the existing folder
        shutil.rmtree(folder_name)
    
    # Create the new folder
    os.makedirs(folder_name)
    
    return folder_name

def slice_file_by_column(df, folder_name, *, column_to_slice_by, columns_to_remove = None):
# Slices the Excel file based on a specific column

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

def compare_dataframes(df1, df2):
# Function to compare DataFrames
    
    # Check if the number of columns is the same
    same_column_count = df1.shape[1] == df2.shape[1]
    
    # Check if column names are the same
    same_columns = set(df1.columns) == set(df2.columns)
    
    # Check if the number of rows is the same
    same_row_count = df1.shape[0] == df2.shape[0]
    
    # Check if the values are the same
    df1_sorted = df1.sort_values(by = list(df1.columns), ascending = True)
    df1_sorted = df1_sorted.reset_index(drop = True)

    df2_sorted = df2.sort_values(by = list(df2.columns), ascending = True)
    df2_sorted = df2_sorted.reset_index(drop = True)
    
    try:
        if df1_sorted.equals(df2_sorted):
            same_values = True
        else:
            same_values = False
    except:
        same_values = False

    # Return results as a dictionary
    return {
        #"file": file,
        "old_report_columns_amount": df1.shape[1],
        "new_report_columns_amount": df2.shape[1],
        "same_column_count": same_column_count,
        "same_columns": same_columns,
        "old_report_row_amount": df1.shape[0],
        "new_report_row_amount": df2.shape[0],
        "same_row_count": same_row_count,
        "same_values": same_values
    }

def combine_files_from_folder(folder_path):
    """
    Combines all Excel files that start with "Extract" from the specified folder into a single Pandas DataFrame.
    
    Parameters:
        folder_path (str): The path to the folder containing the Excel files.
        
    Returns:
        pd.DataFrame: A DataFrame containing the combined data from all matching Excel files.
    """
    all_files = [f for f in os.listdir(folder_path) if f.startswith("Extract") and f.endswith(('.xls', '.xlsx'))]
    
    dataframes = []
    for file in all_files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)
            df["Source File"] = file  # Add a column to track the source file
            dataframes.append(df)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    
    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)
        return combined_df
    else:
        print("No matching Excel files found.")
        return pd.DataFrame()


# ------------------------- Tests below -------------------------

# df1 = pd.read_csv("C:/Users/danie/Downloads/OLD Trad Items with Unverified Price 01022025.csv")
# df2 = pd.read_csv("C:/Users/danie/Downloads/NEW Trad Items with Unverified Price 01022025.csv")

# result = compare_dataframes(df1, df2)
# print(result)