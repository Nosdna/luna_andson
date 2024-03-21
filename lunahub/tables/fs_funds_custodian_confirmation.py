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

class FundsCustodianConfirmation_DownloaderFromLunaHub:
    
    TABLENAME = "fs_funds_custodian_confirmation"
    
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
                    self.status = msg
                    logger.debug(msg)
                
        # Save
        self.df_client = df_client
        self.df_client_fy = df_client_fy
        self.df_client_fy_latest = df_client_fy_latest
        
        return df_client_fy_latest

        

if __name__ == "__main__":
        
    # Test downloader
    if True:
        
        client_number = 50060
        fy = 2023
        uploaddatetime = None
        lunahub_obj = None
        self = FundsCustodianConfirmation_DownloaderFromLunaHub(client_number,
                                                                fy,
                                                                lunahub_obj)
        
        df = self.main()