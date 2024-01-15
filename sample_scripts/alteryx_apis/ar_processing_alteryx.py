# This file is to be run through cmd line.
# Input parameters to be specified through args
# Output will be via a file.

# Import standard package
import os
import sys
import argparse
import importlib.util
import re
import pandas as pd

# Set luna path - Load from settings.py
if True:
    loginid = os.getlogin().lower()
    if loginid == "owghimsiong":
        settings_py_path = r'D:\Desktop\owgs\CODES\luna\settings.py'
    elif loginid == "phuasijia":
        settings_py_path = r'D:\workspace\luna\settings.py'
    else:
        raise Exception (f"Invalid user={loginid}. Please specify the path of settings.py.")
    
    # Import the luna environment through settings.
    # DO NOT TOUCH
    spec = importlib.util.spec_from_file_location("settings", settings_py_path)
    settings = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(settings)

# Import luna packages
import luna
import luna.common as common

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":
    
    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    # parser.add_argument("--client_number", required=True)
    # parser.add_argument("--client_fy", required=True)
    parser.add_argument("--file_fp", required=True)
    parser.add_argument("--file_format", required = True)
    parser.add_argument("--sheet_name", required = True)
    
    # Parse the information
    if True:
        args = parser.parse_args()    
        # client_number = args.client_number
        # fy = args.client_fy
        file_fp = args.file_fp
        file_format = args.file_format
        sheet_name = args.sheet_name

        
    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        # client_number = 40709
        # fy = 2022
        # file_fp = r"D:\workspace\luna\personal_workspace\db\input_file.xlsx"
        file_fp = r"D:\Documents\Project\Internal Projects\20231222 Code integration\MAS forms\Demo\MG\input_file_mg_format1.xlsx"
        file_format = "format1"
        sheet_name = "AR_Aged"
    #############################################

    #process file_fp string
    file_fp = re.findall("r?(.*\.xlsx)|||.*", file_fp)[0]

    self = common.AgedReceivablesReader_Format1(file_fp, sheet_name = sheet_name,
                                                variance_threshold = 0.01)
    
    self.main()

    ar = self.df_processed_long_lcy

    ar[['Left Bin', 'Right Bin']] = ar['Interval (str)'].str.split(' - ', expand=True)
    
    ar['Right Bin'] = ar["Right Bin"].fillna("999999")
    ar['Left Bin'] = ar["Left Bin"].str.replace('+', '')
    ar['Date'] = self.meta_data["Data as at"]

    column_mapper = {'Name'                 : 'NAME',
                     'Left Bin'             : 'LEFTBINVALUE',
                     'Right Bin'            : 'RIGHTBINVALUE',
                     'Currency'             : 'CURRENCY',
                     'Conversion Factor'    : 'CONVERSIONFACTOR',
                     'Value (FCY)'          : 'VALUEFCY',
                     'Value (LCY)'          : 'VALUELCY',
                     'Date'                 : 'DATE'
                     }
    
    ar = ar[column_mapper.keys()]
    
    # Map col names
    ar = ar.rename(columns = column_mapper)
    ar["DATE"] = ar["DATE"].astype(str)

    # Specify temp file
    # output_fn = f"processed_ar_lunahub_{client_number}_{fy}.xlsx"
    output_fn = f"processed_ar_lunahub.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    #output_fp = pyeasylib.check_filepath(output_fp)
    pyeasylib.create_folder_for_filepath(output_fp)    
    ar.to_excel(output_fp, index = False)
    
    print (f"Saved to {output_fp}.")
    