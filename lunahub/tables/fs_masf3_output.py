# Import standard libraries
import pandas as pd
import os
import datetime
import logging

# Import other libraries
import pyeasylib
import luna.lunahub as lunahub

# Configure logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

# class to upload client data
LunaHubBaseUploader = lunahub.LunaHubBaseUploader

class MASForm3Output_LoaderToLunaHub(LunaHubBaseUploader):
    
    COLUMNMAPPER = {
        "var_name"          : "VARNAME",
        "Previous Balance"  : "BALANCEPREVFY",
        "Balance"           : "BALANCE",
        "L/S"               : "LSCODES",
        }
    
    def __init__(self, 
                 df0,
                 client_number, fy, 
                 uploader = None,
                 uploaddatetime = None,
                 lunahub_obj = None):
        '''
        Adds the form 1 output to lunahub.
        
        Input df0:
            - Option 1 with the following columns:
                ['Header 1', 'Header 2', 'Header 3', 'Header 4',
                 'Previous year\n<<<previous_fy>>>\n$',
                 'Current year\n<<<current_fy>>>\n$', 
                 'var_name', 'Has value?', 'L/S',
                 'L/S (intervals)', 'Balance', 'Previous Balance']
                
            - Option 2 with the following columns:
                ["var_name", "Amount", "Subtotal", "Balance", "L/S"]
        '''
        
        # Save as attribute
        self.df0           = df0
        self.client_number = client_number
        self.fy            = fy
        self.tablename     = "fs_masf3_output"

        # check that client number, fy cannot be None
        for name in ["client_number", "fy"]:
            if getattr(self, name) is None:
                raise Exception (f"Attribute {name}=None. Must be specified.")

        # Convert client number to integer
        try:
            self.client_number = int(self.client_number)
        except ValueError as e:
            err = f"{str(e)}\n\nClient number must be a digit."
            raise ValueError (err)
            
        # Init parent class        
        LunaHubBaseUploader.__init__(self,
                                     lunahub_obj    = lunahub_obj,
                                     uploader       = uploader,
                                     uploaddatetime = uploaddatetime,
                                     lunahub_config = None)

    def main(self):
        
        self.process_data()
        
        self.upload_data()
        
        
    def process_data(self):
        
        # df0 may be unfiltered with the following columns
        df0 = self.df0
        
        # Options
        required_columns = list(self.COLUMNMAPPER.keys())
        
        # Filter
        df_processed = df0[required_columns].dropna(subset=['var_name'])
        
        # Do a check that var_name is unique
        pyeasylib.assert_unique(df_processed['var_name'], what='var_name')
        
        # Convert format
        df_processed["L/S"] = df_processed["L/S"].astype(str)
        
        # Map new columns
        df_processed = df_processed.rename(columns=self.COLUMNMAPPER)
        
        # Add client fy
        df_processed["CLIENTNUMBER"] = self.client_number
        df_processed["FY"] = self.fy
        df_processed["UPLOADER"] = self.uploader
        df_processed["UPLOADDATETIME"] = self.uploaddatetime
        
        self.df_processed = df_processed.copy()
        
        return self.df_processed
        
        
    def upload_data(self):
        '''
        This method will do a check before we upload the data to 
        the SQL server.
        '''
            
        # Get current data
        df_processed = self.df_processed.copy()
        
        # Upload
        self.lunahub_obj.insert_dataframe(self.tablename, df_processed)
        

class MASForm3Output_LoaderFromLunaHub:
    
    def __init__(self, client_number, fy, lunahub_obj=None):
        
        
        self.client_number  = int(client_number)
        self.fy             = int(fy)
        self.tablename      = "fs_masf3_output"
        
        if lunahub_obj is None:
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)
        
    def main(self):
        
        df = self.read_from_lunahub()
        
        # Initialise the processed output as None first
        self.df_processed = None
        
        if df is not None:
        
            # filter only necessary cols
            column_mapper = {
                v: k 
                for k, v in MASForm3Output_LoaderToLunaHub.COLUMNMAPPER.items()
                }
            
            df = df[column_mapper.keys()]
            
            # Map col names
            df = df.rename(columns = column_mapper)
            
            # set index
            df = df.set_index("var_name")
            
            self.df_processed = df.copy()
        
        return self.df_processed

    def read_from_lunahub(self):

        # Init output var
        df_client           = None
        df_client_fy        = None
        df_client_fy_latest = None # For any fy, there can be many uploads. 
                                   # Take the latest

        # Read for this client
        query = (
            f"SELECT * FROM {self.tablename} "
            "WHERE "
            f"([CLIENTNUMBER] = {self.client_number})"
            )                
        df_client = self.lunahub_obj.read_table(query = query)
        
        # Check if client data is available
        if df_client.shape[0] == 0:
            msg = f"Data not found for client = {self.client_number}."
            self.status = msg
            logger.warning(msg)
            
        else:
            
            # filter by FY
            df_client_fy = df_client[df_client["FY"]==self.fy]
            
            if df_client_fy.shape[0] == 0:
                msg = (f"Data found for client = {self.client_number} "
                       f"but not for FY = {self.fy}. "
                       f"Available FYs: {df_client['FY'].unique()}")
                self.status = msg
                logger.warning(msg)
            
            else:
                
                # Finally, check the versions
                versions = df_client_fy["UPLOADDATETIME"].unique()
                latest_version = versions.max()
        
                # Filter by latest
                df_client_fy_latest = df_client_fy[
                    df_client_fy["UPLOADDATETIME"] == latest_version]
                
                if len(versions) > 1:
                    msg = (
                        f"Multiple versions for client={self.client_number} "
                        f"and fy={self.fy}: {versions}. "
                        f"Took the latest = {latest_version}."
                        )
                    self.status = msg
                    logger.debug(msg)
                
        # Save
        self.df_client = df_client
        self.df_client_fy = df_client_fy
        self.df_client_fy_latest = df_client_fy_latest
        
        return df_client_fy_latest

if __name__ == "__main__":
    
    
    
    # Uploader
    if False:
        
        # Save the outputdf from form 1 to pickle for testing
        fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\form3\output.df"
        df0 = pd.read_pickle(fp)
        
        client_number = 1
        fy = 2022
        
        self = MASForm3Output_LoaderToLunaHub(df0, client_number, fy)
        self.main()
        
    # Downloader
    if False:
        
        client_number = 1
        fy = 2022
        self = MASForm3Output_LoaderFromLunaHub(client_number, fy)
        df = self.main()
        
            