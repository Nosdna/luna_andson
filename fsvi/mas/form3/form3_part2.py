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
                 df_fp,
                 client_number,
                 fy
                 ):
        
        
        self.outputdf_fp       = df_fp
        self.client_number     = client_number
        self.fy                = fy
        
        self.main()

    def main(self):

        self.read_outputdf()
        # self.process_acct_input()
        self.update_sig_acct()

    def read_outputdf(self):

        self.outputdf = pd.read_excel(self.outputdf_fp)

    # def process_acct_input(self):

    #     df = pd.read_excel(self.sig_acc_output_fp)

    #     new_df = df[df["Declare for current FY?"] == "Yes"]

    #     new_df.fillna("Not provided by user", inplace = True)

    #     self.declared_sig_accts = new_df

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

        sig_acct = sig_accts[sig_accts["Type"] == acct_type]

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
                    self.outputdf.loc[marker+1, "Previous Balance"] = sig_acct_grouped.loc[j, prev_fy]
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
    

if __name__ == "__main__":

    if True:
        fp_dict = {
                # 'sig_acct_output_fp'       : r"D:\Documents\Project\Internal Projects\20231206 Code review\acc_output.xlsx",
                'output_fp'       : r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022_part1.xlsx",
                'final_output_fp' : r"D:\workspace\luna\personal_workspace\tmp\mas_form3_40709_2022.xlsx"
                }

    if True:
        fy = 2022
        client_number = 40709
        outputdf_fp = fp_dict['output_fp']
        # sig_acc_output_fp = fp_dict['sig_acct_output_fp']
        self = MASForm3_Generator_Part2(outputdf_fp, client_number, fy)
    
    # Output to excel
    self.outputdf.to_excel(fp_dict['final_output_fp']) 