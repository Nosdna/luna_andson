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
        self.client_number = client_number
        self.client_name   = client_name
        self.fy_end_date   = fy_end_date
        self.force_insert  = force_insert

        # check that client number, client name, fy end date cannot be None
        for name in ["client_number", "client_name", "fy_end_date"]:
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
                logger.debug("Client info not uploaded as it is already present in lunahub.")
                pass
            
            else:
                
                if self.force_insert:
                    logger.debug("Client info forcely inserted although it is already present in lunahub.")

                    insert()
                    
                else:
                
                    # means we found same client number, but different name or fy info
                    err =  (
                        "Data for client already exists but the info is different.\n\n"
                        f"{existing.T.__repr__()}\n\n"
                        "Please set force_insertion to True to add."
                        )
                    logger.error(err)
                    raise Exception (err)

class ClientInfoLoader_From_LunaHub:

    TABLENAME = "client"
    
    def __init__(self,
                 client_number,
                 lunahub_obj = None):
        
        self.client_number = int(client_number)

        if lunahub_obj is None:
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)
        
        # raise NotImplementedError
            
    def main(self):

        df = self.read_from_lunahub()

        # Initialise the processed output as None first
        self.df_processed = None
        
        if False:
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
        
        return df
    
    def read_from_lunahub(self):

        # Init output var
        df_client           = None
        df_client_latest = None # There can be many uploads. 
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

            # Finally, check the versions
            versions = df_client["UPLOADDATETIME"].unique()
            latest_version = versions.max()
    
            # Filter by latest
            df_client_latest = df_client[
                df_client["UPLOADDATETIME"] == latest_version]
            
            if len(versions) > 1:
                msg = (
                    f"Multiple versions for client={self.client_number} "
                    #f"and fy={self.fy}: {versions}. "
                    f"Took the latest = {latest_version}."
                    )
                self.status = msg
                logger.debug(msg)
                
        # Save
        self.df_client = df_client
        self.df_client_fy_latest = df_client_latest
        
        return df_client_latest
    
    def retrieve_client_info(self, key):

        df = self.main().reset_index()

        dct = df.T.to_dict()[0]

        value = dct.get(key)

        return value
    
    def retrieve_fy_end_date(self, fy):

        fy_end_day      = self.retrieve_client_info("FY_END_DAY")
        fy_end_month    = self.retrieve_client_info("FY_END_MONTH")
                       
        fy_end_date = datetime.datetime(int(fy), int(fy_end_month), int(fy_end_day))

        return fy_end_date
                       

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

    # Test downloader
    if False:
        lunahub_obj = None
        client_number = 71679
        self = ClientInfoLoader_From_LunaHub(client_number, lunahub_obj)

    # Test query
    if True:
        fy = 2022
        client_number = 7167
        lunahub_obj = None
        
        self = ClientInfoLoader_From_LunaHub(client_number, lunahub_obj)
        self.retrieve_fy_end_date(fy)