import luna
import luna.lunahub as lunahub
import pyeasylib

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
        
    
    
    
if __name__ == "__main__":
    
    LUNAHUB_CONFIG = lunahub.LUNAHUB_CONFIG
    
    
    # Connect
    self = LunaHubConnector(**LUNAHUB_CONFIG)