import luna
import luna.lunahub as lunahub
import pyeasylib
import os
import datetime

PyMsSQL = pyeasylib.dblib.PyMsSQL


class LunaHubConnector(PyMsSQL):
    
    def __init__(self, driver, server, database, username, password):
        
        PyMsSQL.__init__(self,
                         driver=driver,
                         server=server,
                         database=database,
                         username=username,
                         password=password)
        
    def future_methods1(self):
        pass
    
    def future_methods2(self):
        pass
        
class LunaHubBaseUploader(PyMsSQL):
    
    def __init__(self, 
                 lunahub_obj    = None,
                 uploader       = None,
                 uploaddatetime = None,
                 lunahub_config = None):
        '''
        Base upload class for LunaHub.
        '''
        
        # Update attr
        self.lunahub_obj    = lunahub_obj
        self.uploader       = uploader
        self.uploaddatetime = uploaddatetime
        self.lunahub_config = None
        
        # Initialise if not provided        
        if self.uploader is None:
            self.uploader = os.getlogin().lower()
        
        if self.uploaddatetime is None:
            self.uploaddatetime = datetime.datetime.now()
               
        if self.lunahub_obj is None:
            
            if self.lunahub_config is None:
                
                self.lunahub_config = lunahub.LUNAHUB_CONFIG
            
            self.lunahub_obj = lunahub.LunaHubConnector(**self.lunahub_config)
            
    
if __name__ == "__main__":
    
    if True:
        
        LUNAHUB_CONFIG = lunahub.LUNAHUB_CONFIG
        
        
        # Connect
        self = LunaHubConnector(**LUNAHUB_CONFIG)