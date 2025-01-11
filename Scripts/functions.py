import csv
from PIL import Image, ImageTk
import pandas as pd
import os
import numpy as np
from pathlib import Path
import sys
import configparser

def convert_to_time_format(hour):
    hours, remainder = divmod(int(hour), 3600)
    minutes, seconds = divmod(remainder, 60)
    return hours, minutes, seconds

def read_csv(csv_file_path):
    """
    Reads a CSV file, skipping the header row, and processes each subsequent row.

    Parameters:
    - csv_file_path (str): The path to the CSV file to be read.

    Returns:
    - A list of rows from the CSV file, where each row is a list of values.
    """
    rows = []  # Initialize an empty list to store rows from the CSV

    with open(csv_file_path, mode='r', encoding='utf-8') as file:
        # Create a CSV reader object
        csv_reader = csv.reader(file)

        # Skip the header row if your CSV has one
        next(csv_reader, None)

        # Loop through each row in the CSV file
        for row in csv_reader:
            rows.append(row)  # Append each row to the list

    return rows  # Return the list of rows

def resize_image(icon_path, size = (15,15)):
    """
    Resizes an image to a fixed size of 15x15 pixels and converts it to a PhotoImage object for Tkinter.

    Parameters:
    - icon_path (str): The file path to the image that needs to be resized.

    Returns:
    - photo (ImageTk.PhotoImage): A Tkinter-compatible PhotoImage object of the resized image.
    """
    img = Image.open(icon_path)
    img = img.resize(size, Image.Resampling.LANCZOS)
    photo = ImageTk.PhotoImage(img)
    return photo

def pandas_read_file(file_path, **kwargs):
        """Determine the Pandas function to read the file based on its extension.
        Args:
        file_path (str): The path to the file to be read.
        **kwargs: Arbitrary keyword arguments that are passed directly to the pandas read function.
        
    Returns:
        A pandas DataFrame or None if an error occurs or the file extension is unsupported.
    """
        file_extension = get_extension(file_path)
        
        # Mapping of file extensions to pandas read functions
        read_functions = {
            '.csv': pd.read_csv,
            '.xls': pd.read_excel,
            '.xlsx': pd.read_excel,
            '.xlsm': pd.read_excel,
            '.json': pd.read_json,
            '.parquet': pd.read_parquet,
        }
        
        if file_extension in read_functions:
            try:
                df = read_functions[file_extension](file_path,**kwargs)
                rows = df.values.tolist()
                return rows
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
        else:
            print(f"Unsupported file extension: {file_extension}")     

def get_extension(file_path):
        """Extract and return the file extension."""
        _, file_extension = os.path.splitext(file_path)
        return file_extension.lower()

def read_excel(excel_file_path):
    """
    Reads an Excel file, skipping the header row, and processes each subsequent row.

    Parameters:
    - excel_file_path (str): The path to the Excel file to be read.

    Returns:
    - A list of rows from the Excel file, where each row is a list of values.
    """
    # Load the Excel file
    df = pd.read_excel(excel_file_path, header=None, skiprows=1, dtype=str)
    df = df.fillna('')

    # Convert DataFrame to a list of rows, each row represented as a list
    rows = df.values.tolist()

    return rows

def create_sap_time_tracker_folder():
    """
    Checks for the existence of a folder named "SAP Time Tracker" in the user's Documents directory
    and creates it if it doesn't exist.
    """
    # Get the path to the user's Documents folder, works for both macOS and Windows
    documents_path = Path.home() / "Documents"

    # Define the full path for the new folder
    new_folder_path = documents_path / "SAP Time Tracker"

    # Check if the folder exists
    if not new_folder_path.exists():
        # If it doesn't exist, create it
        new_folder_path.mkdir()
        print(f"Folder '{new_folder_path}' was created.")
    else:
        print(f"Folder '{new_folder_path}' already exists.")

    new_subfolder_path = documents_path / "SAP Time Tracker" / "Time Records"

    if not new_subfolder_path.exists():
        # If it doesn't exist, create it
        new_subfolder_path.mkdir()
        print(f"Folder '{new_subfolder_path}' was created.")
    else:
        print(f"Folder '{new_subfolder_path}' already exists.")

    file_path = new_folder_path / "chargelines.xlsx"

    # Check if the file exists
    if not file_path.exists():
        # If it doesn't exist, create it using the placeholder function
        create_chargeline_template(file_path)
        print(f"File '{file_path}' was created.")
    else:
        print(f"File '{file_path}' already exists.")

def create_chargeline_template(file_path):
    """
    Creates a blank DataFrame with specified headers and saves it as an Excel file.

    Parameters:
    - file_path (Path or str): The full path, including the file name, where the Excel file will be saved.
    """
    # Define the column headers
    headers = ['Description', 'LDN', 'Rec. Order', 'Network', 'Operation', 'Sub-O']

    # Create a blank DataFrame with these headers
    df = pd.DataFrame(columns=headers)

    # Save the DataFrame as an Excel file at the specified location
    df.to_excel(file_path, index=False)

    print(f"Excel file created at: {file_path}")

def resource_path(relative_path):
    """Get the absolute path to the resource, works for dev and for PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', Path(__file__).parent.parent.absolute())
    return Path(base_path) / relative_path

def is_array_completely_empty(array):
    return all(cell == "" for row in array for cell in row)

def read_config_file():
    pass

def config_autosave():
    pass
    


