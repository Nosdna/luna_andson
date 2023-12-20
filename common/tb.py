# Import standard libraries
import os
import pandas as pd
import numpy as np
import datetime
import logging

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
    

class TBReader_ExcelFormat1:
    
    REQUIRED_HEADERS = ["Account No", "Name", "L/S", "Class"] 
    
    def __init__(self, 
                 fp, sheet_name = 0, 
                 fy_end_date = None,
                 client_number      = None,
                 client_name        = None,
                 insert_to_lunahub  = False):
        '''
        Class to read TB from Excel file.
        
        fy_end_date = the current FY end. Used to verify if the TB
                      for the full year.
                      
        client_number = used to save to the lunahub
        client_name = used to save to the lunahub        
        
        
        Will be processed to the long format.

        Main methods:
            - read_data_from_file
            - process_data
            - get_data_by_fy
        '''
        
        # Initialize file and fy
        self.fp             = fp
        self.sheet_name     = sheet_name
        self.fy_end_date    = fy_end_date  #will validate at the end
        self.client_number  = client_number
        self.client_name    = client_name
        self.insert_to_lunahub = insert_to_lunahub

        # Run
        self.main()
        
    def main(self):
        
        self.read_data_from_file()
        self.process_data()
        
        if self.insert_to_lunahub:
            
            self.load_to_lunahub(client_number = self.client_number, 
                                 client_name = self.client_name)
        
        ####################################################################
        # TO make this consistent across all tb classes
        # Create a tb query class
        tb_query_class = TBQueryClass(self.df_processed_long)
        
        # Unpack the methods to self
        self.get_data_by_fy = tb_query_class.get_data_by_fy
        self.filter_tb_by_fy_and_ls_codes = tb_query_class.filter_tb_by_fy_and_ls_codes
        ##################################################################
        
    def read_data_from_file(self):
        
        if not hasattr(self, 'df1'):

            # Read the main df
            df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols(
                self.fp, sheet_name = self.sheet_name
                )
            
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
            
            # Convert the ls code to interval
            df_processed["L/S (interval)"] = df_processed["L/S"].apply(
                misc.convert_string_to_interval)
            
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
        
        # All other columns
        other_columns = [c for c in df.columns if c not in dates_converted]
        
        # Melt
        df_long = df.melt(
            other_columns, 
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
       
    def load_to_lunahub(self, client_number = None, client_name = None):
        
        # Check that we have clientno
        clientno = self.client_number \
                    if self.client_number is not None \
                    else client_number
        if clientno is None:
            raise Exception ("Client number must be provided during "
                             "initialisation or when calling this method.")
        
        # Check that we have client name
        name = self.client_name \
               if self.client_name is not None \
               else client_name
        if name is None:
            raise Exception ("Client name must be provided during "
                             "initialisation or when calling this method.")
        
        # Upload
        self.tb_uploader_class = TBUploader_To_LunaHub(
            self.df_processed_long,
            clientno,
            name,
            fy_end_date = self.fy_end_date,
            uploader = None,
            uploaddatetime = None,
            lunahub_obj = None)
        self.tb_uploader_class.main()
        
        
        

class TBUploader_To_LunaHub(LunaHubBaseUploader):
    
    def __init__(self,
                 df_processed_long,
                 client_number,
                 client_name,
                 fy_end_date,
                 uploader = None,
                 uploaddatetime = None,
                 lunahub_obj = None):

        self.df_processed_long      = df_processed_long.copy()
        self.client_number          = client_number
        self.client_name            = client_name
        self.fy_end_date            = fy_end_date
        
        # check that client number, client name, fy end date cannot be None
        # Will be checked in the ClientInfoUploader_To_LunaHub class.
        
        # Init parent class        
        LunaHubBaseUploader.__init__(self,
                                     lunahub_obj    = lunahub_obj,
                                     uploader       = uploader,
                                     uploaddatetime = uploaddatetime,
                                     lunahub_config = None)
    
    def main(self):
        
        self.upload_client_info()
        self.upload_tb()
    
    def upload_client_info(self):
        
        # load client table
        client_uploader = tables.client.ClientInfoUploader_To_LunaHub(
            self.client_number, self.client_name, self.fy_end_date, 
            uploader = self.uploader, uploaddatetime = self.uploaddatetime,
            lunahub_obj = self.lunahub_obj,
            force_insert = False
            )
        
        # main
        client_uploader.main()
        
    def upload_tb(self):
        
        # Make a copy
        df = self.df_processed_long.copy()
        
        # Get filter and convert to Lunahub format
        column_mapper = {
            "Account No"    : "ACCOUNTNUMBER",
            "Name"          : "ACCOUNTNAME",
            "L/S"           : "LSCODE",
            "Class"         : "CLASS",
            "Date"          : "DATE",
            "Value"         : "VALUE",
            "FY"            : "FY",
            "Completed FY?" : "COMPLETEDFY"
            }
        
        df = df[list(column_mapper.keys())].rename(columns=column_mapper)
                            
        # Set remaining columns
        df["CLIENTNUMBER"]      = self.client_number
        df["UPLOADER"]          = self.uploader
        df["UPLOADDATETIME"]    = self.uploaddatetime
        df["COMMENTS"]          = None
               
        # --------------------------------------------------
        # load tb table
        self.lunahub_obj.insert_dataframe('tb', df)
        #-------------------------------------------------------
        


class TBLoader_From_LunaHub:
    
    def __init__(self, client_number, fy, uploaddatetime=None):
        '''
        specify uploaddatetime (in str) when there are multiple versions of the same data.
        '''
        
        
        self.client_number  = client_number
        self.fy             = fy
        self.uploaddatetime = uploaddatetime
        
        self.main()
    
    def main(self):
        
        # Load
        df_processed_long = self.load_from_tb()
        
        ####################################################################
        # TO make this consistent across all tb classes
        # Create a tb query class
        tb_query_class = TBQueryClass(self.df_processed_long)
        
        # Unpack the methods to self
        self.get_data_by_fy = tb_query_class.get_data_by_fy
        self.filter_tb_by_fy_and_ls_codes = tb_query_class.filter_tb_by_fy_and_ls_codes
        ##################################################################
        
    def _connect_to_lunahub(self):
        
        if not hasattr(self, 'lunahub_obj'):            
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)
            
        return self.lunahub_obj

    
    def load_from_tb(self):
        
        lunahub_obj = self._connect_to_lunahub()
        
        query = (
            "SELECT * FROM tb "
            "WHERE "
            f"([CLIENTNUMBER] = {self.client_number}) AND (YEAR([DATE]) = {self.fy})"
            )
                
        df = lunahub_obj.read_table(query = query)
        
        # Check if there are multiple records for this run
        version_df = df[["DATE", "UPLOADER", "UPLOADDATETIME", "COMMENTS"]].drop_duplicates()
        
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
            'ACCOUNTNUMBER'     : 'Account No',
            'ACCOUNTNAME'       : 'Name',
            'LSCODE'            : 'L/S',
            'CLASS'             : 'Class',
            'DATE'              : 'Date',
            'VALUE'             : 'Value',
            "FY"                : "FY",
            "COMPLETEDFY"       : "Completed FY?"}
        
        df = df.rename(columns = column_mapper)[list(column_mapper.values())]
        
        # Convert L/S code to intervals
        df["L/S (interval)"] = df["L/S"].astype(str).apply(
            misc.convert_string_to_interval)
            
        self.df_processed_long = df.copy()
        
        return self.df_processed_long
          

class TBQueryClass:
    
    def __init__(self, df_processed_long):
        
        self.df_processed_long = df_processed_long

    def get_data_by_fy(self, fy):
        
        if not hasattr(self, 'gb_fy'):
            
            self.gb_fy = self.df_processed_long.groupby("FY")
        
        # Get
        if fy not in self.gb_fy.groups:
            valid_fys = list(self.gb_fy.groups.keys())
            raise KeyError (f"FY={fy} not found. Valid FYs: {list(valid_fys)}")
            
        return self.gb_fy.get_group(fy)
        
    
    def filter_tb_by_fy_and_ls_codes(self, fy, interval_list):
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
    
    # Test ExcelFormat reader from file format 1
    if False:
        
        # Specify the param fp    
        dirname = os.path.dirname
        luna_fp = dirname(dirname(__file__))
        param_fp = os.path.join(luna_fp, 'templates')
        fp = os.path.join(param_fp, "tb.xlsx")
        
        # Read the tb from file
        sheet_name = "format1"
        fy_end_date = datetime.date(2022, 12, 31)
        client_number = 9999
        client_name = "tester2"
        self = TBReader_ExcelFormat1(fp, 
                                     sheet_name = sheet_name, 
                                     fy_end_date = fy_end_date,
                                     client_number = client_number, 
                                     client_name = client_name,
                                     insert_to_lunahub = False)
        
        df_processed_long = self.df_processed_long
                
        assert False, "End of test."
        
        # Test interval list
        if False:
            
            # We can specify the interval list as pd.Intervals
            interval_list = [
                pd.Interval(7200, 7500, 'both'),
                pd.Interval(3000.1, 3000.1, 'both')
                ]
            
            # Or we can specify as a string of range
            interval_list = ["7200-7500", "3000.1"]
            
            # Filter
            boolean, true_match, false_match = \
                self.filter_tb_by_fy_and_ls_codes(2022, interval_list)
                
            assert False, "End of test."

        
        # Looad and delete from lunahub
        if False:
        
            # Test load to lunahub
            self.load_to_lunahub()
            
            # Delete from luunahub
            self.lunahub_obj.delete('tb', ["CLIENTNUMBER"], pd.DataFrame([[client_number]], columns=["CLIENTNUMBER"]))
            self.lunahub_obj.delete('client', ["CLIENTNUMBER"], pd.DataFrame([[client_number]], columns=["CLIENTNUMBER"]))
            
            
    # Test ExcelFormat reader from lunahub
    if False:
                
        # TESTER TBReader_LunaHub
        client_number = 7167
        fy            = 2022
        
        self = TBLoader_From_LunaHub(client_number, fy)