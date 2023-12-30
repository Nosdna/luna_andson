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
    parser.add_argument("--final_output_fp", required=True)
    
    # Parse the information
    if True:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = int(args.client_fy)
        part1_output_fp = args.part1_output_fp
        final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if False:
        fy                          = 2022
        client_number               = 71679
        credit_quality_output_fp    = rf"D:\workspace\luna\personal_workspace\tmp\mas_f2_{client_number}_{fy}_credit_quality.xlsx"
        # final_output_fp             = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
    #############################################
        
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
        user_response_class = lunahub.tables.fs_masf2_userresponse.MASForm2UserResponse_DownloaderFromLunaHub(
        client_number,
        fy)
        user_inputs = user_response_class.main()  

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
                                    user_inputs = user_inputs
                                    )
    
    # Specify temp file
    output_fn = f"mas_f2_{client_number}_{fy}.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    pyeasylib.create_folder_for_filepath(output_fp)    
    self.write_output(output_fp)

    # Open output file
    if True:
        import webbrowser
        webbrowser.open(output_fp)

    if False:
        # Run and output 
        form3_part2_generator = fsvi.mas.MASForm2_Generator_Part2(
            part1_output_fp, client_number, fy)
        
        form3_part2_generator.write_output(output_fp = final_output_fp)
    
    
    