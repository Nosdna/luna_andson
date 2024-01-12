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
    
    # Parse the information
    if True:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = int(args.client_fy)

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        client_number   = 7167
        fy              = 2022
    #############################################
    
    # Default credit_quality_output fp
    credit_quality_output_fn = f"mas_form2_{client_number}_{fy}_credit_quality.xlsx"
    credit_quality_output_fp = os.path.join(settings.TEMP_FOLDERPATH, credit_quality_output_fn)
    pyeasylib.create_folder_for_filepath(credit_quality_output_fp) 

    if False:
        # Load mapping file
        mas_tb_mapping_fp = os.path.join(settings.LUNA_FOLDERPATH, "parameters", "mas_forms_tb_mapping.xlsx")  
        mapper_class = fsvi.mas.MASTemplateReader_Form3(mas_tb_mapping_fp, sheet_name = "Form 3 - TB mapping")
    
    # Load tb class from LunaHub
    tb_class = common.TBLoader_From_LunaHub(client_number, fy)

    if False:
        # load user response
        user_response_class = lunahub.tables.fs_masf3_userresponse.MASForm3UserResponse_DownloaderFromLunaHub(
            client_number,
            fy)
        user_inputs = user_response_class.main()  
   
    
    self = fsvi.mas.form2.MASForm2_Generator(
        tb_class,
        credit_quality_output_fp,
        fy = fy
        )
    
    # Open output file
    if True:
        import webbrowser
        webbrowser.open(credit_quality_output_fp)

    # # Specify temp file
    # output_fn = f"mas_form2_{client_number}_{fy}_part1.xlsx"
    # output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    # pyeasylib.create_folder_for_filepath(output_fp)    
    # self.write_output(output_fp)
    

