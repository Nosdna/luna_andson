# Import standard libs
import os
import datetime
import sys
import pandas as pd
import re
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
import numpy as np
import logging

# Initialise logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())

# Import luna package and fsvi package

import luna
import luna.common as common
import luna.fsvi as fsvi
from luna.fsvi.mas.template_reader import MASTemplateReader_Form3
import luna.lunahub.tables as tables

import logging

class MASForm3_Generator_Part2:

    def __init__(self,
                 ocr_class, 
                 df_fp,
                 client_number,
                 fy
                 ):
        
        self.ocr_class      = ocr_class
        self.outputdf_fp    = df_fp
        self.client_number  = client_number
        self.fy             = int(fy)
        
        self.main()

    def main(self):

        self.read_outputdf()
        
        self.update_sig_acct()

        self.load_output_to_lunahub()

        self.process_ocr_output()

    def read_outputdf(self):

        self.outputdf = pd.read_excel(self.outputdf_fp)

    def load_sig_accts_from_datahub(self, account_type, fy):
        
        if False:
            # placeholder for db_reader

            # placeholder dataframe
            placeholder_data = {"Account No"    : ["1973558", "1973550"],
                                "Name"          : ["Realised Ex (Gain)/Loss (C)", "Placeholder Account"],
                                "L/S"           : ["7410.4", "7410.4"],
                                "Class"         : ["Revenue - other", "Revenue - other"],
                                "L/S (interval)": [[7410.4, 7410.4], [7410.4, 7410.4]],
                                "Value"         : [10568.0, 500000],
                                "Completed FY?" : [True, True],
                                "Group"         : ["Realised Ex (Gain)/Loss", "Placeholder"],
                                "Type"          : ["Revenue", "Revenue"],
                                "FY"            : [2021, 2021]
                                }
            
            placeholder_df = pd.DataFrame(placeholder_data)

            df = placeholder_df

        # Read from lunahub for current year
        reader_class = tables.fs_masf3_sig_accts.MASForm3SigAccts_DownloaderFromLunaHub(self.client_number, 
                                                                                        fy, lunahub_obj=None)
        reader_class.main()
        df = reader_class.df_processed
        
        # Read from lunahub for previous year
        prev_fy = int(fy) - 1
        reader_class_prevfy = tables.fs_masf3_sig_accts.MASForm3SigAccts_DownloaderFromLunaHub(self.client_number, 
                                                                                        prev_fy, lunahub_obj=None)
        reader_class_prevfy.main()
        df_prevfy = reader_class_prevfy.df_processed

        # Concat for both years
        df_concat = pd.concat([df, df_prevfy], axis=0)
        
        filtered_df = df_concat[df_concat["Type"] == account_type]

        filtered_df["Group"] = filtered_df["Group"].fillna("Not provided by user")

        return filtered_df
    
    def update_sig_acct_by_type(self, account_type):

        if False:
            if account_type.lower() in ["rev", "revenue"]:
                acct_type = "Revenue"
                field = '(l) Other revenue (to specify if significant)'
                prevfy_sig_acct = self.load_sig_accts_from_datahub("Revenue", self.fy-1)
            elif account_type.lower() in ["exp", "expense", "expenses"]:
                acct_type = "Expenses"
                field = "(j) Other expenses (to specify if significant)"
                prevfy_sig_acct = self.load_sig_accts_from_datahub("Expenses", self.fy-1)
            else:
                logger.error(f"Type '{acct_type}' specified is not supported."
                    "Please indicate a different account type.")
            
        if account_type.lower() in ["rev", "revenue"]:
            acct_type = "Revenue"
            field = '(l) Other revenue (to specify if significant)'
            sig_accts = self.load_sig_accts_from_datahub("Revenue", self.fy)
        elif account_type.lower() in ["exp", "expense", "expenses"]:
            acct_type = "Expenses"
            field = "(j) Other expenses (to specify if significant)"
            sig_accts = self.load_sig_accts_from_datahub("Expenses", self.fy)
        else:
            logger.error(f"Type '{acct_type}' specified is not supported."
                  "Please indicate a different account type.")

        sig_acct = sig_accts[sig_accts["Type"] == acct_type].reset_index()

        for i in range(sig_acct.shape[0]):
            if sig_acct.loc[i, 'Group'] == "Not provided by user":
                sig_acct.loc[i, 'Group'] = sig_acct.loc[i, 'Name']
        sig_acct_grouped = pd.pivot_table(sig_acct, values = "Value", index = ["Group"], columns = ["FY"], aggfunc = "sum").reset_index()
        prev_fy = int(self.fy) - 1

        if sig_acct_grouped.empty:
            pass
        else:
            ctr = 0
            for i in range(self.outputdf.shape[0]):
                if self.outputdf.loc[i, 'Header 2'] == field:
                    marker = i
            # while ctr < min(6, sig_acct_grouped.shape[0]):
            #     for j in range(sig_acct_grouped.shape[0]):
            #         self.outputdf.loc[marker+1, 'Header 3'] = sig_acct_grouped.loc[j, 'Group']
            #         self.outputdf.loc[marker+1, "Balance"] = sig_acct_grouped.loc[j, 'Value']
            #         marker += 1
            #         ctr += 1
            for j in range(sig_acct_grouped.shape[0]):
                if ctr < min(6, sig_acct_grouped.shape[0]):
                    self.outputdf.loc[marker+1, 'Header 3'] = sig_acct_grouped.loc[j, 'Group']
                    self.outputdf.loc[marker+1, "Balance"] = sig_acct_grouped.loc[j, self.fy]
                    try:
                        self.outputdf.loc[marker+1, "Previous Balance"] = sig_acct_grouped.loc[j, prev_fy]
                    except:
                        self.outputdf.loc[marker+1, "Previous Balance"] = 0
                    # elif sig_acct_grouped.loc[j, "FY"] == prev_fy:
                    #     print(f"at j: {j} updating for FY{prev_fy}")
                    #     outputdf.loc[marker+1, 'Header 3'] = sig_acct_grouped.loc[j, 'Group']
                    #     outputdf.loc[marker+1, "Previous Balance"] = sig_acct_grouped.loc[j, 'Value']
                    marker += 1
                    ctr += 1
                else:
                    break
        logger.info(f"Significant accounts for {acct_type} updated.")    
    
    def update_sig_acct(self):
        
        self.outputdf = self.outputdf.reset_index(drop = True)

        self.update_sig_acct_by_type("rev")
        self.update_sig_acct_by_type("exp")

    def write_output(self, output_fp = None):
        
        if output_fp is None:
            logger.warning(f"Output not saved as output_fp = {output_fp}.")
        else:
            self.outputdf.to_excel(output_fp)
            logger.info(f"Output saved to {output_fp}.")

    def load_output_to_lunahub(self):

        loader_class = tables.fs_masf3_output.MASForm3Output_LoaderToLunaHub(self.outputdf,
                                                                             self.client_number,
                                                                             self.fy
                                                                             )
        loader_class.main()

    def process_ocr_output(self):
        
        column_mapper = {"var_name"     : "var_name",
                         "previous_fy"  : "Previous Balance",
                         "current_fy"   : "Balance"}

        try:
            ocr_df = self.ocr_class.execute()
        except:
            ocr_df = None
            logger.warning("Unable to process OCR output from Alteryx."
                           "Please check the format of the MAS form provided.")

        if ocr_df is None:
            cols = list(column_mapper.values())
            ocr_df = pd.DataFrame(columns = cols).set_index("var_name")
        else:
            ocr_df = ocr_df[column_mapper.keys()]
            # Map col names
            ocr_df = ocr_df.rename(columns = column_mapper)

        self.ocr_df = ocr_df
        
        return ocr_df
    

if __name__ == "__main__":

    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    print (f"Your luna library is at {luna_folderpath}.")

    fy = 2022
    client_number = 7167

    if True:
        fp_dict = {
                # 'sig_acct_output_fp'       : r"D:\Documents\Project\Internal Projects\20231206 Code review\acc_output.xlsx",
                'output_fp'       : rf"D:\workspace\luna\personal_workspace\tmp\mas_form3_{client_number}_{fy}_part1.xlsx",
                'final_output_fp' : rf"D:\workspace\luna\personal_workspace\tmp\mas_form3_{client_number}_{fy}.xlsx"
                }
    
    if True:
        import importlib.util
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

    if True:
        
        outputdf_fp = fp_dict['output_fp']

        # ocr class
        ocr_fn = f"mas_form3_{client_number}_{fy}_alteryx_ocr.xlsx"
        ocr_fp = os.path.join(luna_folderpath, "personal_workspace", "tmp", ocr_fn)
        ocr_class = fsvi.mas.form3.mas_f3_ocr_output_formatter.OCROutputProcessor(filepath = ocr_fp, sheet_name = "Sheet1", form = "form3", luna_fp = luna_folderpath)


        # sig_acc_output_fp = fp_dict['sig_acct_output_fp']
        self = MASForm3_Generator_Part2(ocr_class, outputdf_fp, client_number, fy)
    
    # Output to excel
    self.outputdf.to_excel(fp_dict['final_output_fp']) 