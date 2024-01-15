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
from luna.fsvi.mas.form2.form2_part2 import MASForm2_Generator_Part2
from luna.fsvi.mas.form2.mas_f2_output_formatter import OutputFormatter

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":

    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    # parser.add_argument("--aic_name", required=True)
    parser.add_argument("--current_qtr", required=True)
    parser.add_argument("--awp_fp", required=False)
    # parser.add_argument("--client_number", required=True)
    # parser.add_argument("--client_fy", required=True)
    # parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    if True:
        args = parser.parse_args()
        # aic_name = args.aic_name
        current_qtr = args.current_qtr
        awp_fp = args.awp_fp
        # client_number = args.client_number
        # fy = int(args.client_fy)
        # part1_output_fp = args.part1_output_fp
        # final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        # fy                          = 2022
        # client_number               = 71679
        # credit_quality_output_fp    = rf"D:\workspace\luna\personal_workspace\tmp\mas_f2_{client_number}_{fy}_credit_quality.xlsx"
        # aic_name                    = "John Smith"
        current_qtr                 = "2022-12-31"
        awp_fp                      = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\f2\MG Based capital calculation Dec 2021-1.xlsx"
        # final_output_fp             = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
    #############################################
    
    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    
    ## Look for sig_account file
    # pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form3_{client_number}_{fy}_sig_accounts.xlsx")
    pattern = os.path.join(settings.TEMP_FOLDERPATH, f"mas_form2_*_*_credit_quality.xlsx")
    list_of_files = glob.glob(pattern)
    cred_quality_fp = max(list_of_files, key=os.path.getctime)
    client_number = int(re.findall("mas_form2_(\d+)_\d{4}_credit_quality.xlsx", cred_quality_fp)[0])
    fy = int(re.findall("mas_form2_\d+_(\d+)_credit_quality.xlsx", cred_quality_fp)[0])

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

    # ocr class
    ocr_fn = f"mas_form2_{client_number}_{fy}_alteryx_ocr.xlsx"
    ocr_fp = os.path.join(settings.LUNA_FOLDERPATH, "personal_workspace", "tmp", ocr_fn)
    ocr_class = fsvi.mas.form2.mas_f2_ocr_output_formatter.OCROutputProcessor(filepath = ocr_fp, sheet_name = "Sheet1", form = "form2", luna_fp = luna_folderpath)


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
    credit_quality_output_fn = f"mas_form2_{client_number}_{fy}_credit_quality.xlsx"
    credit_quality_output_fp = os.path.join(settings.TEMP_FOLDERPATH, credit_quality_output_fn) 

    self = MASForm2_Generator_Part2(tb_class,
                                    mapper_class,
                                    gl_class,
                                    aged_ar_class,
                                    client_class,
                                    ocr_class,
                                    credit_quality_output_fp,
                                    awp_fp,
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

    # Specify OCR output file
    ocr_fn = f"mas_form2_{client_number}_{fy}_ocr.xlsx"
    ocr_fp = os.path.join(settings.TEMP_FOLDERPATH, ocr_fn)
    pyeasylib.create_folder_for_filepath(ocr_fp)    
    self.ocr_df.to_excel(ocr_fp)

    # Initialise client_class
    client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number)

    # Retrieve AIC username
    aic_name = user_response_class.df_client["UPLOADER"].unique()[0]

    # Format output
    formatting_class = OutputFormatter(output_fp, final_output_fp, ocr_fp, client_class, fy, aic_name)

    print (f"Final output saved to {final_output_fp}.")

    # Open output file
    if True:
        import webbrowser
        webbrowser.open(final_output_fp)
        webbrowser.open(self.awp_output_fp)
    
    
    