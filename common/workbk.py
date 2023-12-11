import pandas as pd
import numpy as np
import re

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

class CreditDropdownList:
    ROW = "H2:H1048576"
    
    def __init__(self, filepath):
          self.filepath = filepath
          self.data = None

    def create_dropdown_list(self):

          self.create_wb()
          self.data = self.read_wb()
          return self.data
    
    def create_wb(self):

            wb = openpyxl.load_workbook(self.filepath)
            ws = wb.active

            # Create data validation 
            
            _list = ["Yes", "No"]
            dv = DataValidation(type="list", formula1='"{}"'.format(','.join(_list)))
            
            # Error message
            dv.error ='Your entry is not in the list'
            dv.errorTitle = 'Invalid Entry'

            ws.add_data_validation(dv)

            dv.add(CreditDropdownList.ROW)

            wb.save(self.filepath) 
    
    def read_wb(self):
            
            file = input("File path with credit quality tags: ")
            data = pd.read_excel(file)
            self.data = data
            
            return self.data
            
if __name__ == "__main__":

    workbk_fp = r"D:\gohjiawey\Desktop\Form 3\Credit Quality.xlsx"
    
    workbk = CreditDropdownList(workbk_fp)
    processed_workbk = workbk.create_dropdown_list()
    processed_workbk 