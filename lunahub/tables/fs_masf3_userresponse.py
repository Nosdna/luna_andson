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
    MASForm1UserResponse_UploaderToLunaHub,
    MASForm1UserResponse_DownloaderFromLunaHub
    )

# Form 3 uploader
# Borrow the class from form 1, but change the table name

class MASForm3UserResponse_UploaderToLunaHub(MASForm1UserResponse_UploaderToLunaHub):
    
    TABLENAME = "fs_masf3_userinputs"
    
    def __init__(
        self, 
        user_inputs, client_number, fy_end_date, 
        uploader = None,
        uploaddatetime = None,
        lunahub_obj = None):
        
        MASForm1UserResponse_UploaderToLunaHub.__init__(
            user_inputs, client_number, fy_end_date, 
            uploader = uploader,
            uploaddatetime = uploaddatetime,
            lunahub_obj = lunahub_obj
            )
        





# Form 3 downloader
# Borrow the class from form 1, but change the table name

class MASForm3UserResponse_DownloaderFromLunaHub(MASForm1UserResponse_DownloaderFromLunaHub):
    
    TABLENAME = "fs_masf3_userinputs"
    
    def __init__(self, client_number, fy, lunahub_obj = None):
        
        MASForm1UserResponse_DownloaderFromLunaHub.__init__(
            self, client_number, fy, lunahub_obj = lunahub_obj)
        
    
# Delete class for form 1

if __name__ == "__main__":
    

    # Test downloader
    if True:
        
        client_number = 71679
        fy = 2022
        uploaddatetime = None
        lunahub_obj = None
        self = MASForm3UserResponse_DownloaderFromLunaHub(client_number,
                                                          fy,
                                                          lunahub_obj)
        
       # df = self.main()