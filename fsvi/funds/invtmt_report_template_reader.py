import os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import os

import sys

import luna
import pyeasylib
from luna.common import misc

class FundsInvtmtTemplateReader:

    REQUIRED_HEADERS_SUBLEAD    = ["Current FY", "Previous FY"]
    REQUIRED_HEADERS_PORTFOLIO  = ["Security Name", "ISIN/SEDOL CODE", "Type", "Industry Sector", "Country"
                                  ]
    SHEET_NAME_SUBLEAD          = "<5100-xx>Investment sub-lead"
    SHEET_NAME_PORTFOLIO        = "<5100-xx>Investment Portfolio"

    def __init__(self, fp):

        self.fp = fp

        self.main()

    def main(self):

        self.read_templates()
        self.process_sublead_template()
        self.process_portfolio_template()
        self.get_varname_to_index()

    def read_data_from_file(self, sheet_name, required_headers):

        # Read the main df
        df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols(self.fp,
                                                              sheet_name)
        
        # Strip empty spaces for strings
        df1 = df0.map(lambda s: s.strip() if type(s) is str else s)
        
        # filter out main data
        df1 = pyeasylib.pdlib.get_main_table_from_df(
            df1, required_headers,
            drop_null_rows_for_header_columns = False)
               
                                    
        return df0, df1
    
    def read_templates(self):

        # read sub-lead template
        self.sublead_df0, self.sublead_df1 = self.read_data_from_file(self.SHEET_NAME_SUBLEAD,
                                                                      self.REQUIRED_HEADERS_SUBLEAD
                                                                      )
        
        # read portfolio template
        self.portfolio_df0, self.portfolio_df1 = self.read_data_from_file(self.SHEET_NAME_PORTFOLIO,
                                                                         self.REQUIRED_HEADERS_PORTFOLIO
                                                                         )


    def process_sublead_template(self):

        # get attr
        df1 = self.sublead_df1.copy()

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
        df1["Has value?"] = (df1["Current FY"].notnull() | df1["Previous FY"].notnull())
                   
        # validate that var name is unique
        pyeasylib.assert_no_duplicates(df1["var_name"].dropna().tolist())        
        
        self.sublead_df_processed = df1.copy()

        df0 = self.sublead_df0.copy()

        # get column names
        self.sublead_colname_to_excelcol = pd.Series(
            dict(zip(self.sublead_df_processed.columns, df0.columns))
            )
        self.sublead_colname_to_excelcol.name = "ExcelCol"

    def process_portfolio_template(self):

        # get attr
        df1 = self.portfolio_df1.copy()

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
        df1["Has value?"] = (df1["ISIN/SEDOL CODE"].notnull())
                   
        # # validate that var name is unique
        # pyeasylib.assert_no_duplicates(df1["var_name"].dropna().tolist())        
        
        self.portfolio_df_processed = df1.copy()

        df0 = self.portfolio_df0.copy()

        # get column names
        self.portfolio_colname_to_excelcol = pd.Series(
            dict(zip(self.portfolio_df_processed.columns, df0.columns))
            )
        self.portfolio_colname_to_excelcol.name = "ExcelCol"

    def get_varname_to_index(self):

        df_processed = self.sublead_df_processed.copy()
        df_filtered = df_processed.dropna(subset=["var_name"])
        varname_to_index = df_filtered.reset_index().set_index('var_name')["ExcelRow"]
        self.sublead_varname_to_index = varname_to_index

    

if __name__ == "__main__":

    # Specify the param fp    
    dirname = os.path.dirname
    luna_fp = dirname(dirname(dirname(__file__)))
    param_fp = os.path.join(luna_fp, 'parameters')
    fp = os.path.join(param_fp, "investment_template.xlsx")

    


    self = FundsInvtmtTemplateReader(fp)