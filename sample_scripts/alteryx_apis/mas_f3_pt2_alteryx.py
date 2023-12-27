# This file is to be run through cmd line.
# Input parameters to be specified through args
# Output will be via a file.

# Import standard package
import os
import sys
import argparse
import importlib.util

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

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":

    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    parser.add_argument("--client_number", required=True)
    parser.add_argument("--client_fy", required=True)
    parser.add_argument("--part1_output_fp", required=True)
    parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    if False:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = int(args.client_fy)
        part1_output_fp = args.part1_output_fp
        final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if True:
        fy = 2022
        client_number = 40709
        # sig_acc_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022_sig_accounts.xlsx"
        part1_output_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022_part1.xlsx"
        final_output_fp =     r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
    #############################################

    # Run and output 
    form3_part2_generator = fsvi.mas.MASForm3_Generator_Part2(
        part1_output_fp, client_number, fy)
    
    form3_part2_generator.write_output(output_fp = final_output_fp)
    
    
    