import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
from datetime import datetime
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

import pyeasylib



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
        df_processed["L/S (num)"] = df_processed["L/S"].apply(lambda x: str(x).replace(" ", "").split(","))
    
        self.df_processed = df_processed.copy()
    
    


if __name__ == "__main__":

    fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\parameters\MAS Forms mapping template - Compiled v20231027 (new).xlsx"
    sheet_name = "Form 1 (redesigned)"
    
    self = MASTemplateReader_Form1(fp, sheet_name)
    
    self.read_data_from_file()
    
    self.process_template()
    df_processed = self.df_processed
