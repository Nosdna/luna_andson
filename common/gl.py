# Import standard libraries
import os
import pandas as pd
import numpy as np
import datetime
import logging
import re

# Import other libraries
import pyeasylib
import luna.common.dates as dates
import luna.common.misc as misc
import luna.lunahub as lunahub
import luna.lunahub.tables as tables
LunaHubBaseUploader = lunahub.LunaHubBaseUploader

# Configure logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)


class GLProcessor:        
    def __init__(self, file_path):
        self.file_path = file_path

        self.main()

    def main(self):
        self.promote_row_to_headers()
        self.extract_gl_account_info()
        self.regex_gl_account_no()
        self.forward_fill_gl_account_name()
        self.create_opening_balance_column()
        self.filter_rows()
        self.clean_and_rename_columns()
        self.calculate_amount_column()

        self.reorder_columns()


    def promote_row_to_headers(self):
        gl = pd.read_excel(self.file_path)
        digit = gl[gl.iloc[:,2].notna()].index[0]
        gl    = gl.iloc[digit:,:]
        gl = gl.reset_index(drop = True)
        gl.columns = gl.iloc[0]
        gl = gl[1:]
        nan_columns = gl.columns[gl.columns.isna()]
        gl = gl.drop(columns=nan_columns)

        self.gl = gl.copy()

    def extract_gl_account_info(self):
        gl = self.gl.copy()

        gl['GL Account No'] = gl.iloc[:,0].replace('', np.nan).ffill()
        # self.gl['GL Account Name'] = self.gl['Src'].replace('', np.nan).ffill()
        self.gl = gl

    def regex_gl_account_no(self):
        gl = self.gl.copy()

        regex = r'\d{1}-\d{4}'
        mask = gl['GL Account No'].str.contains(regex, na=False)
        gl['GL Account No'] = gl['GL Account No'].where(mask).fillna(method='ffill')

        self.gl = gl

    def forward_fill_gl_account_name(self):
        gl = self.gl.copy()

        for i in gl.index: 
            if gl.at[i,"ID#"] == gl.at[i,"GL Account No"]:
                gl.at[i,"GL Account Name"] = gl.at[i,"Src"]

        gl["GL Account Name"] = gl["GL Account Name"].ffill()
        # pattern = r'^[A-Z]+\sBank-SGD\s\(\d{3}-\d{6}-\d{1}\)$'
        # self.gl['GL Account Name'] = self.gl['GL Account Name'].astype('str')
        # self.gl['GL Account Name'] = self.gl['GL Account Name'].apply(lambda x: x if re.match(pattern, x) else pd.NA).fillna(method='ffill')

        self.gl = gl 

    def create_opening_balance_column(self):

        gl = self.gl.copy()

        # Extract opening balance
        regex = r'(?i)(\d{0,3},?\d{0,3},?\d+\.?\d{0,2}[c]?)'
        gl['Opening Balance'] = gl['Src'].str.extract(regex)

        for i in gl.index:
            if re.search("(?i)[c]", str(gl.at[i,'Opening Balance'])):
                gl.at[i,'Opening Balance'] = "-" \
                    + gl.at[i,'Opening Balance'][:-1]

        mask = gl['Opening Balance'].notna()
        gl['Opening Balance'] = gl['Opening Balance'].where(mask).ffill()
        gl['Opening Balance'] = pd.to_numeric(
            gl['Opening Balance'].str.replace(',',''), errors='coerce')
        
        self.gl = gl

    def filter_rows(self):
        gl = self.gl.copy()
        gl = gl[gl['ID#'] != 'Beginning Balance:']

        self.gl = gl

    def clean_and_rename_columns(self):
        gl = self.gl.copy()

        gl = gl.drop(gl[gl['Date'].isnull()].index)
        gl = gl.rename(columns={'Date': 'Posting Date', 'ID#': 'Document No', 'Memo': 'Description'})
        gl = gl.drop('Net Activity', axis=1)
        
        self.gl = gl

    def calculate_amount_column(self):
        gl = self.gl.copy()

        gl['Debit'] = pd.to_numeric(gl['Debit'].replace('', np.nan).fillna(0), errors='coerce')
        gl['Credit'] = pd.to_numeric(gl['Credit'].replace('', np.nan).fillna(0), errors='coerce')
        gl['Amount'] = gl['Debit'] - gl['Credit']
        # self.gl['Amount'].fillna("NA", inplace=True)

        self.gl = gl

    def reorder_columns(self):
        column_order = [
            "GL Account No",
            "GL Account Name",
            "Document No",
            "Posting Date",
            "Description",
            "Debit",
            "Credit",
            "Amount",
            "Opening Balance"
        ]
        gl = self.gl.copy()
        gl = gl[column_order]
        self.gl = gl

        return self.gl

class GLLoader_From_LunaHub:

    def __init__(self, client_number, fy, uploaddatetime=None):
        '''
        specify uploaddatetime (in str) when there are multiple versions of the same data.
        '''
        
        
        self.client_number  = int(client_number)
        self.fy             = int(fy)
        self.uploaddatetime = uploaddatetime
        
        self.main()
    
    def main(self):
        
        # Load
        df_processed_long = self.load_from_gl()
        #added by SJ
        self.df_processed_long = df_processed_long
        
        ####################################################################
        # TO make this consistent across all gl classes
        # Create a gl query class
        gl_query_class = GLQueryClass(self.df_processed_long)
        
        # Unpack the methods to self
        self.get_data_by_fy = gl_query_class.get_data_by_fy
        self.filter_gl_by_fy_and_ls_codes = gl_query_class.filter_gl_by_fy_and_ls_codes
        ##################################################################
        
    def _connect_to_lunahub(self):
        
        if not hasattr(self, 'lunahub_obj'):            
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)
            
        return self.lunahub_obj

    
    def load_from_gl(self):
        
        lunahub_obj = self._connect_to_lunahub()
        
        query = (
            "SELECT * FROM gl "
            "WHERE "
            f"([CLIENTNUMBER] = {self.client_number}) AND (YEAR([DATE]) = {self.fy})"
            )
                
        df = lunahub_obj.read_table(query = query)
        
        # Check if there are multiple records for this run
        version_df = df[["DATE", "UPLOADER", "UPLOADDATETIME"]].drop_duplicates()
        
        if (version_df.shape[0] > 1):
            
            if uploaddatetime is None:
                
                msg = f"Multiple records exist.\n\n{version_df.__repr__()}."
                msg += "\n\nPlease set uploaddatetime."
                
                raise Exception (msg)
                
            else:
                
                if isinstance(uploaddatetime, str):
                    uploaddatetime = pd.to_datetime(uploaddatetime)
                
                # Filter
                df = df[df["UPLOADDATETIME"] == uploaddatetime]

        # Map column names
        column_mapper = {
            'ACCOUNTNUMBER'     : 'GL Account No',
            'ACCOUNTNAME'       : 'GL Account Name',
            'DOCUMENTNUMBER'    : 'Document No',
            'POSTINGDATE'       : 'Posting Date',
            'DESCRIPTION'       : 'Description',
            'JETYPE'            : 'JE Type',
            'DEBIT'             : 'Debit',
            'CREDIT'            : 'Credit',
            'AMOUNT'            : 'Amount'}
        
        df = df.rename(columns = column_mapper)[list(column_mapper.values())]
        
        # # Convert L/S code to intervals
        # df["L/S (interval)"] = df["L/S"].astype(str).apply(
        #     misc.convert_string_to_interval)
            
        self.df_processed_long = df.copy()
        self.gl = df.copy()
        
        return self.df_processed_long
    
class GLQueryClass:
    
    def __init__(self, df_processed_long):
        
        self.df_processed_long = df_processed_long

    def get_data_by_fy(self, fy):
        
        if not hasattr(self, 'gb_fy'):
            
            self.gb_fy = self.df_processed_long.groupby("FY")
        
        # Get
        fy = int(fy)
        if fy not in self.gb_fy.groups:
            valid_fys = list(self.gb_fy.groups.keys())
            raise KeyError (f"FY={fy} not found. Valid FYs: {list(valid_fys)}")
            
        return self.gb_fy.get_group(fy)
        
    
    def filter_gl_by_fy_and_ls_codes(self, fy, interval_list):
        '''
        interval_list = a list of pd.Interval
                        a list of strings e.g. ['3', '4-5.5']
        '''
        
        if not isinstance(interval_list, list):
            err = "Input interval_list must be a list of intervals."
            raise Exception (err)
            
        df = self.get_data_by_fy(fy)
        
        # Loop through all the intervals
        temp = []
        for interval in interval_list:
            
            # Convert to interval type, if string is provided
            if type(interval) in [str]:
                interval = misc.convert_string_to_interval(interval)
            
            # Check overlap
            is_overlap = df["L/S (interval)"].apply(lambda i: i.overlaps(interval))
            is_overlap.name = interval
            temp.append(is_overlap)
            
        # Concat
        temp_df = pd.concat(temp, axis=1, names = interval_list)
        
        # final is overlap
        is_overlap = temp_df.any(axis=1)
        
        # get hits
        true_match = df[is_overlap]
        false_match = df[~is_overlap]
        
        return is_overlap, true_match, false_match

    

if __name__ == "__main__":

    if False:
        # GL processing
        gl_fp = r"D:\Documents\Project\Internal Projects\20231222 Code integration\71679_gl.xlsx"
        glp = GLProcessor(gl_fp)
        processed_glp = glp.gl

        processed_glp

    if True:

        client_number   = 40709
        fy              = 2022

        self = GLLoader_From_LunaHub(client_number, fy)

