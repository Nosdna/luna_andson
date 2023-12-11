import luna.lunahub as lunahub
import pandas as pd
import os
import datetime

# class to upload client data
LunaHubBaseUploader = lunahub.LunaHubBaseUploader

class ClientInfoUploader_To_LunaHub(LunaHubBaseUploader):
    
    def __init__(self, 
                 client_number, client_name, fy_end_date, 
                 uploader = None,
                 uploaddatetime = None,
                 lunahub_obj = None,
                 force_insert = False):
        '''
        Adds to the client table in lunahub.
        
        Scenario:
            1) Update when client number is not found in existing db
            2) Do not add when the client is already present, and that
               the name and fy is matched.
            3) Raise an exception when the client is already present, but
               that the name and FY is not matched.
            4) See (3), but data will be uploaded when force_insert=True.
        '''
        
        # Save as attribute
        self.client_number = int(client_number)
        self.client_name   = client_name
        self.fy_end_date   = fy_end_date
        self.force_insert  = force_insert

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
                
        # Convert fyenddate
        if isinstance(self.fy_end_date, str):
            self.fy_end_date = pd.to_datetime(self.fy_end_date)
        
        # Get the day and month
        self.fy_end_day      = self.fy_end_date.day
        self.fy_end_month    = self.fy_end_date.month
        
        
    def upload_data(self):
        '''
        This method will do a check before we upload the data to 
        the SQL server.
        '''
        def insert():
            
            # Prepare df
            df2 = pd.Series(
                [self.client_number, self.client_name,
                 self.fy_end_month, self.fy_end_day,
                 self.uploader, self.uploaddatetime],
                index = ["CLIENTNUMBER", "CLIENTNAME", 
                         "FY_END_MONTH", "FY_END_DAY",
                         "UPLOADER", "UPLOADDATETIME"]).to_frame().T
            
            # Insert
            self.lunahub_obj.insert_dataframe('client', df2)
            
        # Get current data
        df = self.lunahub_obj.read_table("client")
        
        # Check if the data for this client is already present.
        existing = df[df["CLIENTNUMBER"] == self.client_number]
        
        if existing.shape[0] == 0:
            # Not found in db -> Insert
            insert()
        
        else:
        
            # Check if there is any match    
            match_idx = None
            for i in existing.index:
                
                clientnumber = existing.at[i, "CLIENTNUMBER"]
                clientname = existing.at[i, "CLIENTNAME"]
                fyendmonth = existing.at[i, "FY_END_MONTH"]
                fyendday = existing.at[i, "FY_END_DAY"]
            
                if [clientnumber, clientname, fyendmonth, fyendday] == \
                    [self.client_number, self.client_name, self.fy_end_month, self.fy_end_day]:
                    
                    match_idx = i
                    break
            
            if match_idx is not None:
                
                # means a match is found
                # then no need to add
                #print ("exact match found. no need to add")
                pass
            
            else:
                
                if self.force_insert:
                    
                    insert()
                    
                else:
                
                    # means we found same client number, but different name or fy info
                    err =  (
                        "Data for client already exists but the info is different.\n\n"
                        f"{existing.T.__repr__()}\n\n"
                        "Please set force_insertion to True to add."
                        )
                    raise Exception (err)

class ClientInfoLoader_From_LunaHub:
    
    def __init__(self):
        
        raise NotImplementedError
              

if __name__ == "__main__":
    
    # Test adding to table
    if False:
        
        client_number = 1
        client_name = "MA5C"
        fy_end_date = "31 Dec 2023"
        uploader = None
        uploaddatetime = None
        lunahub_obj = None
        force_insert = True

        self = ClientInfoUploader_To_LunaHub(client_number, client_name, fy_end_date,
                                             uploader = uploader,
                                             uploaddatetime = uploaddatetime,
                                             lunahub_obj = lunahub_obj,
                                             force_insert =force_insert)
        
        self.main()
        
        #self.lunahub_obj.delete('client', ["CLIENTNUMBER"], pd.DataFrame([[1]], columns=["CLIENTNUMBER"]))