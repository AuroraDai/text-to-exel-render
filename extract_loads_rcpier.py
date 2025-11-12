"""
Standalone Python script version of Extract Loads RCPier notebook.
This script can be run from command line or imported as a module.
"""

import numpy as np
from pandas import Series, DataFrame
import pandas as pd
import re
import chardet
import sys
import os


def convSPtoDF(x, y):
    """Convert load case to dataframe"""
    lines = [re.split(r'\s{2,}', line.strip()) for line in x.strip().split('\n')]
    df = pd.DataFrame(lines)
    return df


def process_rcpier_file(file_path):
    """
    Process RCPier text file and extract load cases.
    
    Args:
        file_path: Path to the text file
        
    Returns:
        tuple: (df_dict, dframedc, dframell, dframebr, dframews, dframewl)
    """
    # Convert the LPILE txt report to strings
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
        xencode = result['encoding']
    
    print(f"Detected encoding: {xencode}")
    
    with open(file_path, 'r', encoding=xencode) as f:
        textog = f.read()
    
    # Define the begin and end pattern for the re.
    startp = "\n         -------------------------------------------------\n"
    endp = "\n \n      Auto generation details"
    
    # Remove everything after "Selected load groups"
    if "Selected load groups" in textog:
        textog = textog[:textog.find("Selected load groups")]
    
    # Extract table from txt file
    # Main Code
    text = textog
    i = 1
    df_dict = {}
    
    dframedc = pd.DataFrame()
    dframell = pd.DataFrame()
    dframebr = pd.DataFrame()
    dframews = pd.DataFrame()
    dframewl = pd.DataFrame()
    
    loadnameindex = text.find("Loadcase ID:")
    print(f"First loadcase found at index: {loadnameindex}")
    
    while loadnameindex != -1:
        if text[loadnameindex+13] != "W":
            name = text[loadnameindex+13:loadnameindex + 17]
        else:
            name = text[loadnameindex+13:loadnameindex + 45]
            name = name.replace("    Name: ", "-")
            name = name.replace("\n", "")
        
        y = text.find(endp)
        
        # Convert to dframe
        data = text[text.find(startp)+len(startp):y]
        df = convSPtoDF(data, name)
        
        # Delete the dframe when there is 5th column
        if df.shape[1] >= 5:
            # Get the name of the 5th column (index 4, since counting starts from 0)
            col_to_delete = df.columns[4]
            # Drop the column
            df.drop(columns=[col_to_delete], inplace=True)
        
        # Add header
        if df.shape[1] >= 4:
            df.columns = ['Line#', 'Bearing#', 'Direction', 'Loads-Kips']
            
            # Insert Column
            df.insert(0, name, [None] * len(df))
            
            # Put dframe into dictionary
            df_dict[name] = df
            
            # Categorize by load type
            if "DC" in name:
                dframedc = pd.concat([dframedc, df], axis=1)
            elif "WS" in name:
                dframews = pd.concat([dframews, df], axis=1)
            elif "BR" in name:
                dframebr = pd.concat([dframebr, df], axis=1)
            elif "WL" in name:
                dframewl = pd.concat([dframewl, df], axis=1)
            else:
                dframell = pd.concat([dframell, df], axis=1)
        
        l = len(endp)
        text = text[y+l:]
        loadnameindex = text.find("Loadcase ID:")
        i += 1
    
    print(f"Processed {len(df_dict)} load cases")
    print(f"Load case names: {list(df_dict.keys())}")
    
    return df_dict, dframedc, dframell, dframebr, dframews, dframewl


def save_to_excel(file_path, dframedc, dframell, dframebr, dframews, dframewl):
    """Save dataframes to Excel file"""
    excelfilename = file_path.replace(".txt", ".xlsx")
    
    with pd.ExcelWriter(excelfilename, engine='openpyxl') as writer:
        if not dframedc.empty:
            dframedc.to_excel(writer, sheet_name='DC', index=False)
        if not dframell.empty:
            dframell.to_excel(writer, sheet_name='LL', index=False)
        if not dframebr.empty:
            dframebr.to_excel(writer, sheet_name='BR', index=False)
        if not dframews.empty:
            dframews.to_excel(writer, sheet_name='WS', index=False)
        if not dframewl.empty:
            dframewl.to_excel(writer, sheet_name='WL', index=False)
    
    print(f"Excel file saved as: {excelfilename}")
    return excelfilename


if __name__ == "__main__":
    # Command line usage
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Interactive mode - ask for file path
        file_path = input("Enter the path to the text file: ").strip()
    
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found!")
        sys.exit(1)
    
    # Process the file
    df_dict, dframedc, dframell, dframebr, dframews, dframewl = process_rcpier_file(file_path)
    
    # Save to Excel
    excel_file = save_to_excel(file_path, dframedc, dframell, dframebr, dframews, dframewl)
    
    print("\nProcessing complete!")
    print(f"Output file: {excel_file}")

