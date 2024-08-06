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

class FundsFundAdminTxn_DownloaderFromLunaHub:
    
    TABLENAME = "fs_funds_fundadmin_txn"
    
    def __init__(self, 
                 client_number, fy,
                 lunahub_obj = None):
        
        self.client_number  = int(client_number)
        self.fy             = int(fy)
        
        if lunahub_obj is None:
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)

        self.main()

    def main(self):

        df = self.read_from_lunahub()

        # Initialise the processed output as None first
        self.df_processed = None

        if df is not None:
            
            self.df_processed = df.copy()
        
        ####################################################################
        # TO make this consistent across all tb classes
        # Create a txn query class
        txn_query_class = TxnQueryClass(self.df_processed)
        
        # Unpack the methods to self
        self.get_data_by_fy = txn_query_class.get_data_by_fy
        self.filter_txn_by_txntype = txn_query_class.filter_txn_by_txntype
        ##################################################################

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

class TxnQueryClass:
    
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
        
    
    def filter_txn_by_txntype(self, fy, txn_type):
        '''
        interval_list = a list of pd.Interval
                        a list of strings e.g. ['3', '4-5.5']
        '''
        
        df = self.get_data_by_fy(fy)
        
        # Loop through all the txn types, append to temp data
        temp = []

        # Check type match
        is_type = df['TRANSACTIONTYPERSM'].apply(lambda x: x == txn_type)
        is_type.name = txn_type
        temp.append(is_type)
            
        # Concat
        temp_df = pd.concat(temp, axis=1, names = txn_type)
        
        # final is type
        is_type = temp_df.any(axis=1)
        print(is_type)
        # get hits
        true_match = df[is_type]
        false_match = df[~is_type]
        
        return is_type, true_match, false_match        

if __name__ == "__main__":
        
    # Test downloader
    if True:
        
        client_number = 50060
        fy = 2023
        uploaddatetime = None
        lunahub_obj = None
        self = FundsFundAdminTxn_DownloaderFromLunaHub(client_number,
                                                                fy,
                                                                lunahub_obj)
        
        is_type, true_match, false_match = self.filter_txn_by_txntype(2023,'sell')
        print(true_match)
        