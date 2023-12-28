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
from luna.fsvi.mas.form1.mas_f1_output_formatter import OutputFormatter

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":
    
    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    parser.add_argument("--client_number", required=True)
    parser.add_argument("--client_fy", required=True)
    
    # Parse the information
    #args = parser.parse_args()    
    #client_number = args.client_number
    #fy = args.client_fy
        
    client_number = 71679
    fy = 2022

    # Load mapping file
    mas_tb_mapping_fp = os.path.join(settings.LUNA_FOLDERPATH, "parameters", "mas_forms_tb_mapping.xlsx")  
    mapper_class = fsvi.mas.MASTemplateReader_Form1(mas_tb_mapping_fp, sheet_name = "Form 1 - TB mapping")

    # Load tb class from LunaHub
    tb_class = common.TBLoader_From_LunaHub(client_number, fy)
    
    # Load aged ar class
    try:
        aged_ar_class = common.AgedReceivablesLoader_From_LunaHub(client_number, fy)
    except Exception as e:
        if str(e) == f"Data not found for client_number={client_number}.":
            aged_ar_class = None        
        else:
            raise Exception (e)

    # load user response
    user_response_class = lunahub.tables.fs_masf1_userresponse.MASForm1UserResponse_DownloaderFromLunaHub(
        client_number,
        fy)
    user_inputs = user_response_class.main()
    
    self = fsvi.mas.form1.MASForm1_Generator(tb_class, aged_ar_class,
                            mapper_class, fy=fy, fuzzy_match_threshold=80, 
                            user_inputs=user_inputs)
    
    # Specify temp file
    output_fn = f"mas_form1_{client_number}_{fy}.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    #output_fp = pyeasylib.check_filepath(output_fp)
    pyeasylib.create_folder_for_filepath(output_fp)    
    self.outputdf.to_excel(output_fp)
    
    print (f"Saved to {output_fp}.")
    
    # Run OutputProcessor
    template_fn = r"parameters\mas_forms_tb_mapping.xlsx"
    template_fp = os.path.join(settings.LUNA_FOLDERPATH, template_fn)
    final_output_fn = f"mas_form1_formatted_{client_number}_{fy}.xlsx"
    final_output_fp = os.path.join(settings.TEMP_FOLDERPATH, final_output_fn)
    formatting_class = OutputFormatter(output_fp, final_output_fp)

    print (f"Final output saved to {final_output_fp}.")
    