import re
import pandas as pd
import os
import numpy as np
from pathlib import Path
from win32com import client
from striprtf.striprtf import rtf_to_text

def find_decimal_followed_by_keyword(text, keyword):
    """
    Finds decimal numbers followed by a specific keyword or phrase in the text.
    """
    pattern = rf"(?i)(-?\d+\.\d+ | \.\d+)(\s*(?:u\/|\w+\s+|\w+\s\w+\s+)?{keyword})"
    return re.findall(pattern, text)

def find_profile_directions(text):
    """
    Finds specific words indicating profile directions in the text.
    """
    pattern = r"(?i)surplus|scant|positive|negative"
    return re.findall(pattern, text)

def find_defect_extent(text, direction_words, over_keyword):
    """
    Finds decimal numbers followed by 'over' keyword, considering specific direction words.
    """
    if direction_words:
        pattern = rf"(?i)(-?\d+\.\d+|\.\d+)(\s*(?:\w+\s*)?{over_keyword})"
        return re.findall(pattern, text)
    return []

def print_defect_results(values, measure_unit, direction=None):
    """
    Prints the defect results, including maximum defect extent and direction if provided.
    """
    if values:
        max_value = max(values)
        if direction:
            print(f"Max extent: {max_value} in {direction} direction")
            print("********************************************************")
        else:
            print(f"Max extent: {max_value} {measure_unit}")
            print("********************************************************")
            
def process_defects(defect_text):
    """Takes in Defect Text and looks for defect data"""
    # Process max and min defects
    res_max = find_decimal_followed_by_keyword(defect_text, "Max")
    res_min = find_decimal_followed_by_keyword(defect_text, "Min")

    max_values = []
    min_values = []

    if res_max:
        for val in res_max:
            if type(val) is tuple:
                max_values.append(abs(float(val[0])))
                print("Defect Extent: ", abs(float(val[0]))," o/max")
        print("********************************************************")
    
    if res_min:
        for val in res_min:
            if type(val) is tuple:
                min_values.append(abs(float(val[0])))
                print("Defect Extent: ", abs(float(val[0])), ' u/min')
                # print(float(val[0]))
        print("********************************************************")


    # Process profile direction defects
    profile_directions = find_profile_directions(defect_text)
    res_profile = find_defect_extent(defect_text, profile_directions, "over")
    
    max_profile_scant = []
    max_profile_surplus = []

    if res_profile:
        for index, val in enumerate(res_profile,0):
            if type(val) is tuple and (profile_directions[index].lower() in ("scant", "negative")):
                max_profile_scant.append(abs(float(val[0])))
                print("Defect Extent: ", abs(float(val[0])), "in scant direction")

            elif type(val) is tuple and (profile_directions[index].lower() in ("surplus"  "positive")):
                max_profile_surplus.append(abs(float(val[0])))
                print("Defect Extent: ", abs(float(val[0])), "in surplus direction")
    


    # Process visual defects
    res_vis = re.findall(r"(?i)VIS", defect_text)

    # Print results
    print_defect_results(max_values, "o/max")
    print_defect_results(min_values, "u/min")
    print_defect_results(max_profile_scant, "over", "scant")
    print_defect_results(max_profile_surplus, "over", "surplus")

    if res_vis:
        print("This is a visual defect")
        print("********************************************************")

def get_extension(file_path):
        """Extract and return the file extension."""
        _, file_extension = os.path.splitext(file_path)
        return file_extension.lower()

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
                return read_functions[file_extension](file_path,**kwargs)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
        else:
            print(f"Unsupported file extension: {file_extension}")     

def read_rich_text_file(path, delete = True):
    with open(path) as infile:
        content = infile.read()
        text = rtf_to_text(content)
        # print(text)
        
    # Deletes Temp File
    if delete:   
        os.remove(path)
    return text
    
# test = "Your test string here is 0.004 over SCANT 0.001 over surplus"
# test = "This is 1.50 over max and 1.20 under min"
# process_defects(test)
            
# path= Path(r"C:\Users\M337199\Documents\Copy of bn_qn_proc_tool_v1.1.3.xlsm")

# data=pandas_read_file(path)
# print(data)
    
# read_rich_text_file(r"C:\Users\M337199\Documents\SAP\SAP GUI\temp_dlt.rtf")