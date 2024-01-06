# This file is to be run through cmd line.
# Input parameters to be specified through args
# Output will be via a file.

# Import standard package
import os
import sys
import argparse
import importlib.util
import glob
import re

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
import luna.fsvi as fsvi
import luna.lunahub as lunahub
from luna.fsvi.mas.form3.mas_f3_output_formatter import OutputFormatter

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":

    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    parser.add_argument("--aic_name", required=True)
    parser.add_argument("--mic_name", required=True)
    # parser.add_argument("--client_number", required=True)
    # parser.add_argument("--client_fy", required=True)
    # parser.add_argument("--part1_output_fp", required=True)
    # parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    if True:
        args = parser.parse_args()
        aic_name = args.aic_name
        mic_name = args.mic_name
        # client_number = args.client_number
        # fy = int(args.client_fy)
        # part1_output_fp = args.part1_output_fp
        # final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        fy              = 2022
        client_number   = 40709
        aic_name        = "John Smith"
        mic_name        = "Jane Doe"
        # sig_acc_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022_sig_accounts.xlsx"
        # part1_output_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022_part1.xlsx"
        # final_output_fp =     r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
    #############################################

    ## Look for sig_account file
    # pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form3_{client_number}_{fy}_sig_accounts.xlsx")
    pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form3_*_*_sig_accounts.xlsx")
    list_of_files = glob.glob(pattern)
    sig_acc_fp = max(list_of_files, key=os.path.getctime)
    client_number = re.findall("mas_form3_(\d+)_\d{4}_sig_accounts.xlsx", sig_acc_fp)[0]
    fy = re.findall("mas_form3_\d+_(\d{4})_sig_accounts.xlsx", sig_acc_fp)[0]
        
    part1_output_fn = f'mas_form3_{client_number}_{fy}_part1.xlsx'
    output_fn = f'mas_form3_{client_number}_{fy}.xlsx'
    part1_output_fp = os.path.join(settings.TEMP_FOLDERPATH, part1_output_fn)
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
        
    # Run and output 
    self = fsvi.mas.MASForm3_Generator_Part2(
        part1_output_fp, client_number, fy)
    
    self.write_output(output_fp = output_fp)

    # Initialise client_class
    client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number)

    # Run OutputProcessor
    template_fn = r"parameters\mas_forms_tb_mapping.xlsx"
    template_fp = os.path.join(settings.LUNA_FOLDERPATH, template_fn)
    final_output_fn = f"mas_form3_formatted_{client_number}_{fy}.xlsx"
    final_output_fp = os.path.join(settings.TEMP_FOLDERPATH, final_output_fn)
    formatting_class = OutputFormatter(output_fp, final_output_fp, fy,
                                       client_class, aic_name, mic_name)
    
    
    # Open output file
    if True:
        import webbrowser
        webbrowser.open(final_output_fp)