import pandas as pd
import numpy as np
import re


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
            "Opening Balance",
        ]
        gl = self.gl.copy()
        gl = gl[column_order]
        self.gl = gl

        return self.gl


    

if __name__ == "__main__":
    # GL processing
    gl_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\f2\GL FY2023.xlsx"
    glp = GLProcessor(gl_fp)
    processed_glp = glp.gl

    processed_glp

