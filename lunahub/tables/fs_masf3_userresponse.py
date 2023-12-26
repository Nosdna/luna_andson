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
from fs_masf1_userresponse import (
    MASForm1UserResponse_UploaderToLunaHub,
    MASForm1UserResponse_DownloaderFromLunaHub
    )

# Alias
# Form 3 uploader
MASForm3UserResponse_UploaderToLunaHub = MASForm1UserResponse_UploaderToLunaHub
MASForm3UserResponse_UploaderToLunaHub.TABLENAME = 'fs_masf3_userinputs'

# Form 3 downloader
MASForm3UserResponse_DownloaderFromLunaHub = MASForm1UserResponse_DownloaderFromLunaHub
MASForm3UserResponse_DownloaderFromLunaHub.TABLENAME = 'fs_masf3_userinputs'


if __name__ == "__main__":
    

    # Test downloader
    if True:
        
        client_number = 7167
        fy = 2022
        uploaddatetime = None
        lunahub_obj = None
        self = MASForm1UserResponse_DownloaderFromLunaHub(client_number,
                                                          fy,
                                                          
                                                          lunahub_obj)
        
        df = self.main()