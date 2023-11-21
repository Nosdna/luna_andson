import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
import datetime
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import pyeasylib
import dates

class TBReader_ExcelFormat1:
    
    REQUIRED_HEADERS = ["Account No", "Name", "L/S", "Class"] 
    
    def __init__(self, fp, sheet_name = 0, fy_end_date = None):
        '''

        Main methods:
            - read_data_from_file
            - process_data
            - get_data_by_fy
        '''
        
        # Initialize file and fy
        self.fp             = fp
        self.sheet_name     = sheet_name
        self.fy_end_date    = fy_end_date  #will validate at the end       

        # Run
        self.main()
        
    def main(self):
        
        self.read_data_from_file()
        self.process_data()
        
        
    def read_data_from_file(self):
        
        if not hasattr(self, 'df1'):

            # Read the main df
            df0 = pd.read_excel(self.fp, sheet_name = self.sheet_name, 
                                engine = 'openpyxl')
            
            # Strip empty spaces for strings
            df_processed = df0.applymap(lambda s: s.strip() if type(s) is str else s)
            
            # filter out main data
            df_processed = pyeasylib.pdlib.get_main_table_from_df(
                df_processed, self.REQUIRED_HEADERS)
            
            # Save as attr
            self.df0 = df0
            self.df_processed = df_processed
                                    
        return self.df_processed


    def process_data(self):
        '''
        # Processing
        1) convert all date format to datetime.date
        2) convert string column
        3) convert value to float
        4) convert to long format
        '''

        if not hasattr(self, "df_long"):
            
            # Read the data
            df_processed = self.read_data_from_file()
                                
            # Get the date columns
            date_values = [c 
                            for c in df_processed.columns 
                            if c not in self.REQUIRED_HEADERS]
            
            # validate date columns and map to date
            date_to_converted = self._process_dates(date_values)
            dates_converted = list(date_to_converted.values())
            
            # Rename the date column
            df_processed = df_processed.rename(columns = date_to_converted)
            
            # Format the data types
            for c in self.REQUIRED_HEADERS:
                df_processed[c] = df_processed[c].astype(str)
                
            for c in dates_converted:
                df_processed[c] = df_processed[c].astype(float)
            
            # Convert to long
            df_processed_long = self._convert_to_long_format(df_processed, dates_converted)
            
            # Classify FY
            df_processed_long, fy_class = self._classify_fy(df_processed_long)
            
            # Save attributes
            self.df_processed       = df_processed
            self.df_processed_long  = df_processed_long
            self.date_to_converted  = date_to_converted
            self.dates_converted    = dates_converted
            self.fy_class           = fy_class
            
        return self.df_processed_long
    
    def get_data_by_fy(self, fy):
        
        if not hasattr(self, 'gb_fy'):
            
            self.gb_fy = self.df_processed_long.groupby("FY")
        
        # Get
        if fy not in self.gb_fy.groups:
            valid_fys = list(self.gb_fy.groups.keys())
            raise KeyError (f"FY={fy} not found. Valid FYs: {list(valid_fys)}")
            
        return self.gb_fy.get_group(fy)
        
        
    def _process_dates(self, date_values):
        
        # Validate that date is of correct type
        for dt in date_values:
            dates.is_valid_types(
                dt, dates.DATE_TYPES, 
                raise_on_invalid = True
                )
        
        # Then convert to date obj
        # those already in date format does not have d.date() method.
        date_to_converted = {
            d : d.date() if hasattr(d, "date") else d 
            for d in date_values
            }
        
        return date_to_converted
    
    
    def _convert_to_long_format(self, df, dates_converted):
        
        # Melt
        df_long = df.melt(
            self.REQUIRED_HEADERS, 
            dates_converted,
            var_name = "Date",
            value_name = "Value"
            )
                
        return df_long
        
    def _classify_fy(self, df_processed_long):
    
        # Make a local copy
        df_processed_long = df_processed_long.copy()
        
        # Call a FY generator class
        fy_class = dates.FYGenerator(fy_end_date = self.fy_end_date)
        
        ##### ADD THE FY ########
        # note: slower if we do it in one line as there are many duplicate dates
        #    df_processed_long["FY"] =  df_processed_long["Date"].apply(fy_class.get_date_fy)

        # Extract all the dates
        date_values = df_processed_long["Date"].unique()
        date_val_to_fy = {d: fy_class.get_date_fy(d) for d in date_values}

        # Add back to the df
        df_processed_long["FY"] = df_processed_long["Date"].map(date_val_to_fy)
        
        ##### Check if it's full year of data #######
        date_val_to_completeness = {
            d: d == fy_class.get_fy_dates(date_val_to_fy[d], which='end')
            for d in date_values
            }
        df_processed_long["Completed FY?"] = df_processed_long["Date"].map(date_val_to_completeness)
        
        return df_processed_long.copy(), fy_class
        
       


if __name__ == "__main__":
    
    
    if False:
        
        # Read the tb
        fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\sample_tb_data.xlsx"
        sheet_name = "format1"
        fy_end_date = datetime.date(2022, 12, 31)
        
        self = TBReader_ExcelFormat1(fp, sheet_name = sheet_name, fy_end_date = fy_end_date)
        
        
        
    
        assert False
        file = input("Enter the TB file path: ")
        # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\myer gold\Myer Gold Investment Management - 2022 TB.xlsx
        fy = int(input("Enter the Financial Year (e.g. 2022): "))
        tb = TBReader(file, fy)
        # df = tb.process_icm()     # only use for ICM
        df = tb.tbprocessing()
        
        