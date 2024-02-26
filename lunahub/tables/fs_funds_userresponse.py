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

# Load form 1 user response
from luna.lunahub.tables.fs_masf1_userresponse import (
    MASForm1UserResponse_DownloaderFromLunaHub
    )

# Funds downloader
# Borrow the class from form 1, but change the table name

class FundsUserResponse_DownloaderFromLunaHub(MASForm1UserResponse_DownloaderFromLunaHub):
    
    TABLENAME = "fs_funds_userinputs"
    
    def __init__(self, client_number, fy, lunahub_obj = None):
        
        MASForm1UserResponse_DownloaderFromLunaHub.__init__(
            self, client_number, fy, lunahub_obj = lunahub_obj)
        

if __name__ == "__main__":
    

    # Test downloader
    if True:
        
        client_number = 50060
        fy = 2023
        uploaddatetime = None
        lunahub_obj = None
        self = FundsUserResponse_DownloaderFromLunaHub(client_number,
                                                       fy,
                                                       lunahub_obj)
        
       # df = self.main()