# This file is to be run through cmd line.
# Input parameters to be specified through args
# Output will be via a file.

# Import standard package
import os
import sys
import argparse
import importlib.util
import time

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
    if False:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = int(args.client_fy)

    #############################################
    ## FOR DEBUGGING ONLY ##
    if True:
        client_number = 7167
        fy = 2022
    #############################################
    
    # Default output fp
    sig_acc_output_fp = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form3_{client_number}_{fy}_sig_accounts.xlsx")
    
    # Load mapping file
    mas_tb_mapping_fp = os.path.join(settings.LUNA_FOLDERPATH, "parameters", "mas_forms_tb_mapping.xlsx")  
    mapper_class = fsvi.mas.MASTemplateReader_Form3(mas_tb_mapping_fp, sheet_name = "Form 3 - TB mapping")

    ######################################################################
    # CURRENT YEAR
    ######################################################################
    
    # Load tb class from LunaHub
    tb_class = common.TBLoader_From_LunaHub(client_number, fy)

    # load user response
    for attempt in range(12):
        time.sleep(5)
        user_response_class = lunahub.tables.fs_masf3_userresponse.MASForm3UserResponse_DownloaderFromLunaHub(
            client_number,
            fy)
        user_inputs = user_response_class.main()
        if user_inputs is not None:
            break
        elif user_inputs is None and attempt == 11:
            raise Exception (f"Data not found for specified client {client_number} or FY {fy}.")
        else:
            continue
   
    
    # Current fy    
    current_fy_class = fsvi.mas.MASForm3_Generator(
        tb_class,
        mapper_class,
        sig_acc_output_fp,
        client_number,
        fy = fy,
        user_inputs = user_inputs
        )
    
    #####################################################################
    # PREVIOUS YEAR
    #####################################################################
    prevfy = fy-1
    
    # Load tb class from LunaHub
    tb_class_prevfy = common.TBLoader_From_LunaHub(client_number, prevfy)

    # load user response
    for attempt in range(12):
            time.sleep(5)
            user_response_class_prevfy = lunahub.tables.fs_masf3_userresponse.MASForm3UserResponse_DownloaderFromLunaHub(
                client_number,
                prevfy)
            user_inputs_prevfy = user_response_class_prevfy.main()
            if user_inputs is not None:
                break
            elif user_inputs is None and attempt == 11:
                raise Exception (f"Data not found for specified client {client_number} or FY {fy}.")
            else:
                continue

    try:  
        # Previous fy    
        prevfy_class = fsvi.mas.MASForm3_Generator(
            tb_class_prevfy,
            mapper_class,
            None,
            client_number,
            fy = prevfy,
            user_inputs = user_inputs_prevfy
            )
        
        # Append prev year data to current year
        current_fy_class.outputdf['Previous Balance'] = prevfy_class.outputdf["Balance"]

    except:
        # Append prev year data to current year
        current_fy_class.outputdf['Previous Balance'] = None

    
    
    # Specify temp file
    output_fn = f"mas_form3_{client_number}_{fy}_part1.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    pyeasylib.create_folder_for_filepath(output_fp)    
    current_fy_class.write_output(output_fp)

    # Open output file
    if True:
        import webbrowser
        webbrowser.open(sig_acc_output_fp)
    

