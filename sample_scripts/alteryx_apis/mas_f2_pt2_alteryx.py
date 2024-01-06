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
from luna.fsvi.mas.form2.mas_f2_output_formatter import OutputFormatter

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":

    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    parser.add_argument("--aic_name", required=True)
    parser.add_argument("--mic_name", required=True)
    parser.add_argument("--current_qtr", required=True)
    # parser.add_argument("--client_number", required=True)
    # parser.add_argument("--client_fy", required=True)
    # parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    if True:
        args = parser.parse_args()
        aic_name = args.aic_name
        mic_name = args.mic_name
        current_qtr = args.current_qtr
        # client_number = args.client_number
        # fy = int(args.client_fy)
        # part1_output_fp = args.part1_output_fp
        # final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        fy                          = 2022
        client_number               = 71679
        credit_quality_output_fp    = rf"D:\workspace\luna\personal_workspace\tmp\mas_f2_{client_number}_{fy}_credit_quality.xlsx"
        aic_name                    = "John Smith"
        mic_name                    = "Jane Doe"
        current_qtr                 = "31/12/2022"
        # final_output_fp             = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
    #############################################
    
    ## Look for sig_account file
    # pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form3_{client_number}_{fy}_sig_accounts.xlsx")
    pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_f2_*_*_credit_quality.xlsx")
    list_of_files = glob.glob(pattern)
    cred_quality_fp = max(list_of_files, key=os.path.getctime)
    client_number = re.findall("mas_f2_(\d+)_\d{4}_credit_quality.xlsx", cred_quality_fp)[0]
    fy = re.findall("mas_f2_\d+_(\d+)_credit_quality.xlsx", cred_quality_fp)[0]

    # Load AR from LunaHub
    if True:
        aged_ar_class = common.AgedReceivablesLoader_From_LunaHub(client_number, fy)
        
    # Load from LunaHub
    if True:
        tb_class = common.TBLoader_From_LunaHub(client_number, fy)
        
    # Form 2 mapping
    if False:
        
        mas_tb_mapping_fp = os.path.join(luna_folderpath, "parameters", "mas_forms_tb_mapping.xlsx")
        print (f"Your mas_tb_mapping_fp is at {mas_tb_mapping_fp}.")
        
        # Load the class
        mapper_class = fsvi.mas.MASTemplateReader_Form1(mas_tb_mapping_fp, sheet_name = "Form 2 - TB mapping")
    
        # process df is here:
        df_processed = mapper_class.df_processed  # need to build methods

    if True:
        # Load mapping file
        mas_tb_mapping_fp = os.path.join(settings.LUNA_FOLDERPATH, "parameters", "mas_forms_tb_mapping.xlsx")  
        mapper_class = fsvi.mas.MASTemplateReader_Form1(mas_tb_mapping_fp, sheet_name = "Form 2 - TB mapping")
    
    # Load GL from LunaHub
    if True:
        gl_class = common.gl.GLLoader_From_LunaHub(client_number, fy)

    # Retrieve FY end date
    if True:
        client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number, None)

    # Load user input from LunaHub
    if True:

        for attempt in range(12):
            time.sleep(5)
            user_response_class = lunahub.tables.fs_masf2_userresponse.MASForm2UserResponse_DownloaderFromLunaHub(
                client_number,
                fy)
            user_inputs = user_response_class.main()
            if user_inputs is not None:
                break
            elif user_inputs is None and attempt == 11:
                raise Exception (f"Data not found for specified client {client_number} or FY {fy}.")
            else:
                continue

    # CLASS
        
    # Credit quality output fp
    # output_folderpath = rf"D:\workspace\luna\personal_workspace\tmp"
    credit_quality_output_fn = f"mas_f2_{client_number}_{fy}_credit_quality.xlsx"
    credit_quality_output_fp = os.path.join(settings.TEMP_FOLDERPATH, credit_quality_output_fn) 
    
    self = fsvi.mas.MASForm2_Generator_Part2(tb_class,
                                    mapper_class,
                                    gl_class,
                                    aged_ar_class,
                                    client_class,
                                    credit_quality_output_fp,
                                    settings.TEMP_FOLDERPATH,
                                    client_number,
                                    fy,
                                    current_qtr,
                                    user_inputs = user_inputs
                                    )
    
    # Specify temp file
    output_fn = f"mas_form2_{client_number}_{fy}.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    pyeasylib.create_folder_for_filepath(output_fp)    
    self.write_output(output_fp)

    final_output_fn = f"mas_form2_formatted_{client_number}_{fy}.xlsx"
    final_output_fp = os.path.join(settings.TEMP_FOLDERPATH, final_output_fn)

    # Initialise client_class
    client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number)

    # Format output
    formatting_class = OutputFormatter(output_fp, final_output_fp, client_class, aic_name, mic_name)

    print (f"Final output saved to {final_output_fp}.")

    # Open output file
    if True:
        import webbrowser
        webbrowser.open(final_output_fp)
    
    
    