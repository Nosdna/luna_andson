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

import logging

class MASForm3_Generator_Part2:

    def __init__(self, 
                 df_fp,
                 sig_acc_output_fp,
                 fy,
                 ):
        
        
        self.outputdf_fp       = df_fp
        self.sig_acc_output_fp = sig_acc_output_fp
        self.fy                = fy
        
        self.main()

    def main(self):

        self.read_outputdf()
        self.process_acct_input()
        self.update_sig_acct()

    def read_outputdf(self):

        self.outputdf = pd.read_excel(self.outputdf_fp)

    def process_acct_input(self):

        df = pd.read_excel(self.sig_acc_output_fp)

        new_df = df[df["Declare for current FY?"] == "Yes"]

        new_df.fillna("Not provided by user", inplace = True)

        self.declared_sig_accts = new_df

    def load_sig_accts_from_datahub(self, account_type, fy):

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

        if df.empty:
            pass
        else:
            query = f"Type  == '{account_type}' and FY == {fy}"
            df = df.query(query)
            df["Indicator"] = "Declared in prev FY"

        return df
    
    def update_sig_acct_by_type(self, account_type, df):

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

        sig_acct = df[df["Type"] == acct_type]

        for i in range(sig_acct.shape[0]):
            if sig_acct.loc[i, 'Group'] == "Not provided by user":
                sig_acct.loc[i, 'Group'] = sig_acct.loc[i, 'Name']
                
        sig_acct_grouped = sig_acct.groupby(["Group"]).agg({"Value" : "sum"}).reset_index()

        sig_acct["Account No"] = sig_acct["Account No"].astype(str)
        cols_to_keep = ["Account No", "Group"]
        sig_acct = sig_acct[cols_to_keep]
        prevfy_sig_acct["Account No"] = prevfy_sig_acct["Account No"].astype(str)

        prevfy_sig_acct = prevfy_sig_acct.drop(columns = ["Group", "Type"])
        combined = prevfy_sig_acct.merge(sig_acct,
                                         how      = "right",
                                         on       = "Account No",
                                         suffixes = ("_prev", "_curr")
                                         ).reset_index()
        
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
                    self.outputdf.loc[marker+1, "Balance"] = sig_acct_grouped.loc[j, 'Value']
                    marker += 1
                    ctr += 1
                else:
                    break

        if sig_acct_grouped.empty:
            pass
        else:
            ctr = 0
            for i in range(self.outputdf.shape[0]):
                if self.outputdf.loc[i, 'Header 2'] == field:
                    marker = i
            # while ctr < min(6, combined.shape[0]):
            #     for j in range(combined.shape[0]):
            #         if self.outputdf.loc[marker+1, "Header 3"] == combined.loc[j, "Group"]:
            #             self.outputdf.loc[marker+1, 'Header 3'] = combined.loc[j, 'Group']
            #             self.outputdf.loc[marker+1, "Previous Balance"] = combined.loc[j, 'Value']
            #             marker += 1
            #             ctr += 1
            for j in range(combined.shape[0]):
                if ctr < min(6, combined.shape[0]):
                    if self.outputdf.loc[marker+1, "Header 3"] == combined.loc[j, "Group"]:
                        self.outputdf.loc[marker+1, 'Header 3'] = combined.loc[j, 'Group']
                        self.outputdf.loc[marker+1, "Previous Balance"] = combined.loc[j, 'Value']
                        marker += 1
                        ctr += 1
                else:
                    break
        
    
    def update_sig_acct(self):
        
        self.outputdf = self.outputdf.reset_index(drop = True)

        self.update_sig_acct_by_type("rev", self.declared_sig_accts)
        self.update_sig_acct_by_type("exp", self.declared_sig_accts)
        
    def write_output(self, output_fp = None):
        
        if output_fp is None:
            logger.warning(f"Output not saved as output_fp = {output_fp}.")
        else:
            self.outputdf.to_excel(output_fp)
            logger.info(f"Output saved to {output_fp}.")
    

if __name__ == "__main__":

    fp_dict = {
               'sig_acct_output_fp'       : r"D:\Documents\Project\Internal Projects\20231206 Code review\acc_output.xlsx",
               'output_fp'                : r"D:\Documents\Project\Internal Projects\20231206 Code review\form_3_output.xlsx",
               'final_output_fp'          : r"D:\Documents\Project\Internal Projects\20231206 Code review\form_3_output_test.xlsx"
               }

    if True:
        fy = 2022
        outputdf_fp = fp_dict['output_fp']
        sig_acc_output_fp = fp_dict['sig_acct_output_fp']
        gen_part_2 = MASForm3_Generator_Part2(outputdf_fp, sig_acc_output_fp, fy)
    
    # Output to excel
    gen_part_2.outputdf.to_excel(fp_dict['final_output_fp']) 