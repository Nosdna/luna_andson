# Import standard libraries
import pandas as pd
import os
import datetime
import logging

# Import other libraries
import luna.lunahub as lunahub

# Configure logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())
logger.setLevel(logging.DEBUG)

# class to upload client data
LunaHubBaseUploader = lunahub.LunaHubBaseUploader

class MASForm1UserResponse_UploaderToLunaHub(LunaHubBaseUploader):
    
    COLUMN_MAPPER = {
        "Index"       : "VARNAME",
        "Question"    : "PROMPT",
        "Answer"      : "USERRESPONSE"}
    
    TABLENAME = "fs_masf1_userinputs"
    
    def __init__(self, 
                 user_inputs, client_number, fy_end_date, 
                 uploader = None,
                 uploaddatetime = None,
                 lunahub_obj = None):
        '''
        '''
        
        # Save as attribute
        self.user_inputs   = user_inputs
        self.client_number = client_number
        self.fy_end_date   = fy_end_date
        
        # Init parent class        
        LunaHubBaseUploader.__init__(self,
                                     lunahub_obj    = lunahub_obj,
                                     uploader       = uploader,
                                     uploaddatetime = uploaddatetime,
                                     lunahub_config = None)
    
    def _process(self):
        
        # Make a copy
        user_inputs_luna = self.user_inputs.copy().reset_index()
        
        # Map columns
        user_inputs_luna = user_inputs_luna.rename(columns=self.COLUMN_MAPPER)

        # Convert to string
        user_inputs_luna = user_inputs_luna.astype(str)
        
        # Add meta data
        user_inputs_luna["CLIENTNUMBER"]    = self.client_number
        user_inputs_luna["FY"]              = self.fy_end_date.year
        user_inputs_luna["UPLOADER"]        = self.uploader
        user_inputs_luna["UPLOADDATETIME"]  = self.uploaddatetime
        user_inputs_luna["COMMENTS"]        = None
    
        # Save and return
        self.user_inputs_processed = user_inputs_luna.copy()
        
        return self.user_inputs_processed
        
    def upload_to_lunahub(self):
        
        # Get the data with mapped columns
        user_inputs_processed = self._process()
        
        # Upload
        self.lunahub_obj.insert_dataframe(self.TABLENAME, user_inputs_processed)


class MASForm1UserResponse_DownloaderFromLunaHub:
    
    TABLENAME = MASForm1UserResponse_UploaderToLunaHub.TABLENAME
    
    def __init__(self, 
                 client_number, fy,
                 lunahub_obj = None):
        
        self.client_number  = int(client_number)
        self.fy             = int(fy)
        
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
                for k, v in MASForm1UserResponse_UploaderToLunaHub.COLUMN_MAPPER.items()
                }
            
            df = df[column_mapper.keys()]
            
            # Map col names
            df = df.rename(columns = column_mapper)
            
            # set index
            df = df.set_index("Index")
            
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
            f"SELECT * FROM {self.TABLENAME} "
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
                    logger.debug(msg)
                
        # Save
        self.df_client = df_client
        self.df_client_fy = df_client_fy
        self.df_client_fy_latest = df_client_fy_latest
        
        return df_client_fy_latest
    
        


if __name__ == "__main__":
    
    # Test uploader
    if False:
        user_inputs = pd.DataFrame({'Question': {'trade_debt_fund_mgmt': 'List of client names related to fund management (trade debtors): ',
          'total_trade_cred': 'Total trade creditors amount: $',
          'trade_cred_fund_mgmt': 'Trade creditors for fund managment amount: $',
          'amount_due_to_director': 'Enter the client account numbers for amounts due to director or connected persons: ',
          'loans_from_related_co': 'Enter the client account numbers for loans from related company or associated persons: ',
          'amount_due_from_director_secured': 'Enter the client account numbers for amounts due from director and connected persons (secured): ',
          'amount_due_from_director_unsecured': 'Enter the client account numbers for amounts due from director and connected persons (unsecured): ',
          'loans_to_related_co': 'Enter the client account numbers for loans to related company or associated person: '},
         'Answer': {'trade_debt_fund_mgmt': 'NA',
          'total_trade_cred': '2902',
          'trade_cred_fund_mgmt': '0',
          'amount_due_to_director': 'NA',
          'loans_from_related_co': 'NA',
          'amount_due_from_director_secured': 'NA',
          'amount_due_from_director_unsecured': 'NA',
          'loans_to_related_co': 'NA'}})
        user_inputs.index.name = "Index"
        
        client_number = 99999
        fy_end_date = pd.to_datetime("31-Dec 2022")
        
        self = MASForm1UserResponse_UploaderToLunaHub(user_inputs, client_number, 
                                                      fy_end_date)
        
        self.upload_to_lunahub()
        
    # Test downloader
    if False:
        
        client_number = 7167
        fy = 2022
        uploaddatetime = None
        lunahub_obj = None
        self = MASForm1UserResponse_DownloaderFromLunaHub(client_number,
                                                          fy,
                                                          
                                                          lunahub_obj)
        
        df = self.main()