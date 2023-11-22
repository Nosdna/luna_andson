import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
from datetime import datetime
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

import pyeasylib

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
        df1["Has value?"] = (df1["Amount"].notnull() | df1["Subtotal"].notnull())
                   
        # validate that var name is unique
        pyeasylib.assert_no_duplicates(df1["var_name"].dropna().tolist())        
        
        self.df_processed = df1.copy()
        
        #
        self._format_leadsheet_code()
        
        
    def _format_leadsheet_code(self):
        
        # Get attr
        df_processed = self.df_processed.copy()
    
        # Create ls - combine both amt and subtotal
        df_processed["L/S"] = df_processed["Subtotal"].fillna(df_processed["Amount"])
    
        # split
        df_processed["L/S (intervals)"] = df_processed["L/S"].apply(lambda x: str(x).replace(" ", "").split(","))
    
        self.df_processed = df_processed.copy()
    
    def get_ls_codes_by_varname(self, varname):

        
        if not hasattr(self, 'varname_to_lscodes'):
            
            # only process if the varname to lscodes is not defined yet
            
            # Get the data, filter and set index
            df_processed = self.df_processed.copy()
            df_filtered = df_processed.dropna(subset=["var_name"])
            varname_to_lscodes = df_filtered.set_index('var_name')["L/S (intervals)"]
            
            # check that no duplicated varname
            pyeasylib.assert_no_duplicates(varname_to_lscodes.index)
        
            # convert to intervals
            varname_to_lscodes = varname_to_lscodes.apply(misc.convert_list_of_string_to_interval)
        
            # save as attr
            self.varname_to_lscodes = varname_to_lscodes
                
        # Get 
        return self.varname_to_lscodes.at[varname]


if __name__ == "__main__":

    fp = r"D:\Desktop\owgs\CODES\luna\parameters\mas_forms_tb_mapping.xlsx"
    sheet_name = "Form 1 - TB mapping"
    
    self = MASTemplateReader_Form1(fp, sheet_name)
    
    self.read_data_from_file()
    
    self.process_template()
    df_processed = self.df_processed
    
    self.get_ls_codes_by_varname("puc_rev_reserve")
    
