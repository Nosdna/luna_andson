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
    if False:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = int(args.client_fy)
        part1_output_fp = args.part1_output_fp
        final_output_fp = args.final_output_fp

    #############################################
    ## FOR DEBUGGING ONLY ##
    if True:
        fy                          = 2022
        client_number               = 71679
        credit_quality_output_fp    = r"D:\workspace\luna\personal_workspace\tmp\mas_form2_71679_2022_credit_quality.xlsx"
        final_output_fp             = r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
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
    folderpath = rf"D:\workspace\luna\personal_workspace\tmp\mas_f2_{client_number}_{fy}_"
    credit_quality_output_fn = "credit_quality.xlsx"
    credit_quality_output_fp = folderpath + credit_quality_output_fn

    output_fn = "output.xlsx"
    output_fp = folderpath + output_fn

    output_wp_fn = "output_wp.xlsx"
    output_wp_fp = folderpath + output_wp_fn
    
    # Output fp

    self = fsvi.mas.MASForm2_Generator_Part2(tb_class,
                                    mapper_class,
                                    gl_class,
                                    aged_ar_class,
                                    client_class,
                                    credit_quality_output_fp,
                                    client_number,
                                    fy,
                                    user_inputs = user_inputs
                                    )
    
    # Get df by varname
    # filtered_tb = self.filter_tb_by_varname('current_asset_trade_debt_other')
    
    # Output to excel 
    # self.outputdf.to_excel("draftf2.xlsx") 


    
    if False:
        # Run and output 
        form3_part2_generator = fsvi.mas.MASForm2_Generator_Part2(
            part1_output_fp, client_number, fy)
        
        form3_part2_generator.write_output(output_fp = final_output_fp)
    
    
    