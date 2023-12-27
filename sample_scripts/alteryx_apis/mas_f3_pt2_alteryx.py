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
    parser.add_argument("--client_fy", required=True)
    parser.add_argument("--sig_acc_fp", required=True)
    parser.add_argument("--part1_output_fp", required=True)
    parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    args = parser.parse_args()    
    fy = int(args.client_fy)
    sig_acc_fp = args.sig_acc_fp
    part1_output_fp = args.part1_output_fp
    final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        fy = 2022
        sig_acc_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\tmp\mas_form3_40709_2022_sig_accounts.xlsx"
        part1_output_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
        final_output_fp =     r"D:\Desktop\owgs\CODES\luna\personal_workspace\tmp\mas_form3_40709_2022_final.xlsx"
    #############################################

    # Run and output 
    form3_part2_generator = fsvi.mas.MASForm3_Generator_Part2(
        part1_output_fp, sig_acc_fp, fy)
    
    form3_part2_generator.write_output(output_fp = final_output_fp)
    
    
    