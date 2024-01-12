import os
import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
from datetime import datetime
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import sys
sys.path.append("D:\gohjiawey\Desktop\Form 3\CODES")

import luna
import pyeasylib
from luna.common import misc 
# import luna.commmon.misc as misc 

import luna.common.misc as misc

class MASTemplateReader_Form1:
    
    REQUIRED_HEADERS = ["Amount", "Subtotal"] 

    def __init__(self, fp, sheet_name):
        
        self.fp = fp
        self.sheet_name = sheet_name
        
        self.main()
    
    def main(self):
        
        self.read_data_from_file()
        self.process_template()
        self.get_varname_to_ls_codes()
        
    def read_data_from_file(self):
        
        # Read the main df
        df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols(self.fp,
                                                              self.sheet_name)
        
        # Strip empty spaces for strings
        df1 = df0.applymap(lambda s: s.strip() if type(s) is str else s)
        
        # filter out main data
        df1 = pyeasylib.pdlib.get_main_table_from_df(
            df1, self.REQUIRED_HEADERS,
            drop_null_rows_for_header_columns = False)
               
        # Save as attr
        self.df0 = df0
        self.df1 = df1
                                    
        return self.df1

        
    def process_template(self):
        
        # Get attr
        df1 = self.df1.copy()
        
        # new cols
        cols = []
        count = 1
        for c in df1.columns:
            if pd.isnull(c):
                cols.append(f"Header {count}")
                count += 1
            else:
                cols.append(c)
        df1.columns = cols
                
        # tag if has values
        df1["Has value?"] = (df1["Amount"].notnull() | df1["Subtotal"].notnull())
                   
        # validate that var name is unique
        pyeasylib.assert_no_duplicates(df1["var_name"].dropna().tolist())        
        
        self.df_processed = df1.copy()
        
        #
        self._format_leadsheet_code()
        
        # Get column names
        self.colname_to_excelcol = pd.Series(
            dict(zip(self.df_processed.columns, self.df0.columns))
            )
        self.colname_to_excelcol.name = "ExcelCol"
    
        
        
    def _format_leadsheet_code(self):
        
        # Get attr
        df_processed = self.df_processed.copy()
    
        # Create ls - combine both amt and subtotal
        df_processed["L/S"] = df_processed["Subtotal"].fillna(df_processed["Amount"])
    
        # split
        df_processed["L/S (intervals)"] = df_processed["L/S"].apply(lambda x: str(x).replace(" ", "").split(","))
    
        self.df_processed = df_processed.copy()
        
        
    def get_varname_to_ls_codes(self):

        if not hasattr(self, 'varname_to_lscodes'):
            
            # only process if the varname to lscodes is not defined yet
            
            # Get the data, filter and set index
            df_processed = self.df_processed.copy()
            df_filtered = df_processed.dropna(subset=["var_name"])
            varname_to_lscodes = df_filtered.set_index('var_name')["L/S (intervals)"]
            varname_to_index = df_filtered.reset_index().set_index('var_name')["ExcelRow"]
            
            # check that no duplicated varname
            pyeasylib.assert_no_duplicates(varname_to_lscodes.index)

            # NOTE: edited by SJ to accommodate formulas on mapping template 20240111
            varname_to_lscodes_temp = varname_to_lscodes.copy()
            varname_to_lscodes_temp = varname_to_lscodes_temp.to_frame()
            # varname_to_lscodes_temp["L/S (intervals)"] = varname_to_lscodes_temp["L/S (intervals)"].astype(str)
            substr = "="
            varname_to_lscodes_ls = varname_to_lscodes_temp[~varname_to_lscodes_temp["L/S (intervals)"].astype(str).apply(lambda x: any(substr in item for item in x))]
            varname_to_lscodes_formula = varname_to_lscodes_temp[varname_to_lscodes_temp["L/S (intervals)"].astype(str).apply(lambda x: any(substr in item for item in x))]
            substr2 = "<"
            varname_to_lscodes_ls = varname_to_lscodes_ls[~varname_to_lscodes_ls["L/S (intervals)"].astype(str).apply(lambda x: any(substr2 in item for item in x))]
            varname_to_lscodes_f1 = varname_to_lscodes_ls[varname_to_lscodes_temp["L/S (intervals)"].astype(str).apply(lambda x: any(substr2 in item for item in x))]

            varname_to_lscodes_ls = varname_to_lscodes_ls["L/S (intervals)"].squeeze()
            varname_to_lscodes_formula = varname_to_lscodes_formula["L/S (intervals)"].squeeze()
            varname_to_lscodes_f1 = varname_to_lscodes_f1["L/S (intervals)"].squeeze()
        


            if False:
                varname_to_lscodes_temp = varname_to_lscodes.copy().astype(str)
                pattern = "^['=.*']$"
                varname_to_lscodes_ls = varname_to_lscodes[~varname_to_lscodes_temp.str.contains(pattern)]
                varname_to_lscodes_formula = varname_to_lscodes[varname_to_lscodes_temp.str.contains(pattern)]
                pattern2 = ".*<<<.*>>>.*"
                varname_to_lscodes_ls = varname_to_lscodes_ls[~varname_to_lscodes_temp.str.contains(pattern2)]
                varname_to_lscodes_f1 = varname_to_lscodes[varname_to_lscodes_temp.str.contains(pattern2)]


            # convert to intervals
            varname_to_lscodes_ls = varname_to_lscodes_ls.apply(misc.convert_list_of_string_to_interval)

            # varname_to_lscodes = varname_to_lscodes.append(varname_to_lscodes_formula)
            varname_to_lscodes = pd.concat([varname_to_lscodes_ls, varname_to_lscodes_formula, varname_to_lscodes_f1], axis = 0)

            # ## ORIGINAL ##
            # # convert to intervals
            # varname_to_lscodes = varname_to_lscodes.apply(misc.convert_list_of_string_to_interval)
            # ## ORIGINAL - END ##
        
            # save as attr
            self.varname_to_lscodes = varname_to_lscodes
            self.varname_to_index   = varname_to_index
        
        return self.varname_to_lscodes
    
    def get_ls_codes_by_varname(self, varname):
        
        s = self.get_varname_to_ls_codes()
        
        if varname not in s.index:
            
            raise KeyError (f"Input varname={varname} not found.")
        
        return s.at[varname]
    
    def get_varname_to_formula(self):

        full_frame = self.varname_to_lscodes.copy()
        full_frame_temp = full_frame.copy().astype(str)
        pattern = "^\['=.*"
        formula_frame = full_frame[full_frame_temp.str.contains(pattern)]

        return formula_frame


class MASTemplateReader_Form3:
    
    REQUIRED_HEADERS = ["Previous year\n<<<previous_fy>>>\n$",
                        "Current year\n<<<current_fy>>>\n$"]

    def __init__(self, fp, sheet_name):
        
        self.fp = fp
        self.sheet_name = sheet_name

        self.main()
        
    def main(self):
        
        self.read_data_from_file()
        self.process_template()
        self.get_varname_to_ls_codes()

        
    def read_data_from_file(self):
        
        # Read the main df
        df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols(self.fp,
                                                              self.sheet_name)

        
        # Strip empty spaces for strings
        df1 = df0.applymap(lambda s: s.strip() if type(s) is str else s)
        
        # filter out main data
        df1 = pyeasylib.pdlib.get_main_table_from_df(
            df1, self.REQUIRED_HEADERS,
            drop_null_rows_for_header_columns = False)
               
        # Save as attr
        self.df0 = df0
        self.df1 = df1
                                    
        return self.df1
    
    def process_template(self):
            
            # Get attr
            df1 = self.df1.copy()
            
            # new cols
            cols = []
            count = 1
            for c in df1.columns:
                if pd.isnull(c):
                    cols.append(f"Header {count}")
                    count += 1
                else:
                    cols.append(c)
            df1.columns = cols
            
            # ffill
            if False:
                cols_to_ffill = df1.columns[:df1.columns.tolist().index("Amount")]
                for c in cols_to_ffill:
                    df1[c] = df1[c].ffill()
                
            # tag if has values
            df1["Has value?"] = (df1["Previous year\n<<<previous_fy>>>\n$"].notnull() | df1["Current year\n<<<current_fy>>>\n$"].notnull())
                    
            # validate that var name is unique
            pyeasylib.assert_no_duplicates(df1["var_name"].dropna().tolist())        
            
            self.df_processed = df1.copy()
            
            #
            self._format_leadsheet_code()
            
        
            # Get column names
            self.colname_to_excelcol = pd.Series(
                dict(zip(self.df_processed.columns, self.df0.columns))
                )
            self.colname_to_excelcol.name = "ExcelCol"
        
        


    def _format_leadsheet_code(self):
        
        # Get attr
        df_processed = self.df_processed.copy()
    
        # Create ls - combine both amt and subtotal
        df_processed["L/S"] = df_processed["Previous year\n<<<previous_fy>>>\n$"].fillna(df_processed["Current year\n<<<current_fy>>>\n$"])
    
        # split
        df_processed["L/S (intervals)"] = df_processed["L/S"].apply(lambda x: str(x).replace(" ", "").split(","))
    
        self.df_processed = df_processed.copy()


    def get_varname_to_ls_codes(self):

        if not hasattr(self, 'varname_to_lscodes'):
            
            # only process if the varname to lscodes is not defined yet
            
            # Get the data, filter and set index
            df_processed = self.df_processed.copy()
            df_filtered = df_processed.dropna(subset=["var_name"])
            varname_to_lscodes = df_filtered.set_index('var_name')["L/S (intervals)"]
            varname_to_index = df_filtered.reset_index().set_index('var_name')["ExcelRow"]
            
            # check that no duplicated varname
            pyeasylib.assert_no_duplicates(varname_to_lscodes.index)
        
            # drop those with [nan]
            varname_to_lscodes = varname_to_lscodes[varname_to_lscodes.apply(lambda s:s!= ['nan'])]

            # NOTE: edited by SJ to accommodate formulas on mapping template 20240111
            # varname_to_lscodes_temp = varname_to_lscodes.copy().astype(str)
            # pattern = "^=.*"
            # varname_to_lscodes_ls = varname_to_lscodes[~varname_to_lscodes_temp.str.contains(pattern)]
            # varname_to_lscodes_formula = varname_to_lscodes[varname_to_lscodes_temp.str.contains(pattern)]

            varname_to_lscodes_temp = varname_to_lscodes.copy()
            varname_to_lscodes_temp = varname_to_lscodes_temp.to_frame()
            # varname_to_lscodes_temp["L/S (intervals)"] = varname_to_lscodes_temp["L/S (intervals)"].astype(str)
            substr = "="
            varname_to_lscodes_ls = varname_to_lscodes_temp[~varname_to_lscodes_temp["L/S (intervals)"].astype(str).apply(lambda x: any(substr in item for item in x))]
            varname_to_lscodes_formula = varname_to_lscodes_temp[varname_to_lscodes_temp["L/S (intervals)"].astype(str).apply(lambda x: any(substr in item for item in x))]
           
            varname_to_lscodes_ls = varname_to_lscodes_ls["L/S (intervals)"].squeeze()
            varname_to_lscodes_formula = varname_to_lscodes_formula["L/S (intervals)"].squeeze()

            # convert to intervals
            varname_to_lscodes_ls = varname_to_lscodes_ls.apply(misc.convert_list_of_string_to_interval)

            # varname_to_lscodes = varname_to_lscodes.append(varname_to_lscodes_formula)
            varname_to_lscodes = pd.concat([varname_to_lscodes_ls, varname_to_lscodes_formula], axis = 0)

            # ## ORIGINAL ##
            # # convert to intervals
            # varname_to_lscodes = varname_to_lscodes.apply(misc.convert_list_of_string_to_interval)
            # ## ORIGINAL - END ##

            # save as attr
            self.varname_to_lscodes = varname_to_lscodes
            self.varname_to_index   = varname_to_index
        

        return self.varname_to_lscodes
    
    def get_ls_codes_by_varname(self, varname):
        
        s = self.get_varname_to_ls_codes()
        
        if varname not in s.index:
            
            raise KeyError (f"Input varname={varname} not found.")
        
        return s.at[varname]
    
    def get_varname_to_formula(self):

        full_frame = self.varname_to_lscodes.copy()
        full_frame_temp = full_frame.copy().astype(str)
        pattern = ".*=.*"
        formula_frame = full_frame[full_frame_temp.str.contains(pattern)]

        return formula_frame

    # def get_ls_codes_by_varname(self, varname):    
        
    #     if not hasattr(self, 'varname_to_lscodes'):
            
    #         # only process if the varname to lscodes is not defined yet

    #         # Get the data, filter and set index
    #         df_processed = self.df_processed.copy()
    #         df_filtered = df_processed.dropna(subset=["var_name"])
    #         varname_to_lscodes = df_filtered.set_index('var_name')["L/S (num)"]
            
    #         # check that no duplicated varname
    #         pyeasylib.assert_no_duplicates(varname_to_lscodes.index)
 
    #         # save as attr
    #         self.varname_to_lscodes = varname_to_lscodes
                
    #     # Get 
    #     return self.varname_to_lscodes.at[varname]

        s = self.get_varname_to_ls_codes()

        if varname not in s.index:
            
            raise KeyError (f"Input varname={varname} not found.")
        
        return s.at[varname]
 
if __name__ == "__main__":

    # Specify the param fp    
    dirname = os.path.dirname
    luna_fp = dirname(dirname(dirname(__file__)))
    param_fp = os.path.join(luna_fp, 'parameters')
    fp = os.path.join(param_fp, "mas_forms_tb_mapping.xlsx")
    sheet_name = "Form 2 - TB mapping"
    
    
    # Main
    self = MASTemplateReader_Form1(fp, sheet_name)
    
    if False:
        #fp = r"D:\Desktop\owgs\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"
        fp  = r"D:\gohjiawey\Desktop\Form 3\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"

        sheet_name = "Form 3 - TB mapping"

        self = MASTemplateReader_Form3(fp, sheet_name)

        self.read_data_from_file()
        
        self.process_template()
        df_processed = self.df_processed

        self.get_ls_codes_by_varname("exp_prov_dtf_debts")
        #df_processed.to_excel("testing.xlsx")
        
        df_processed

    if False:
        fp  = r"D:\gohjiawey\Desktop\Form 3\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"
        sheet_name = "Form 2 - TB mapping"
    
        self.get_ls_codes_by_varname("puc_rev_reserve")


    if False:
        #fp = r"D:\Desktop\owgs\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"
        fp  = r"D:\gohjiawey\Desktop\Form 3\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"

        sheet_name = "Form 3 - TB mapping"

        self = MASTemplateReader_Form3(fp, sheet_name)

        self.read_data_from_file()
        
        self.process_template()
        df_processed = self.df_processed

        self.get_ls_codes_by_varname("exp_prov_dtf_debts")
        #df_processed.to_excel("testing.xlsx")
        
    
