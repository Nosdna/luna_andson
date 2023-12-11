import pandas as pd
import numpy as np
import pyeasylib
import re
import os
import datetime
import luna.lunahub as lunahub
import luna.lunahub.tables as tables
import numpy as np

AGED_AR_TO_LUNAHUB_MAPPER = {
    "Name"              : "NAME",
    "Currency"          : "CURRENCY",
    "Conversion Factor" : "CONVERSIONFACTOR",
    "Interval (str)"    : "INTERVALSTR",
    "Interval"          : "INTERVAL",
    "Value (FCY)"       : "VALUEFCY",
    "Value (LCY)"       : "VALUELCY",
    "FY"                : "FY",
    "Date"              : "DATE"
    }


LunaHubBaseUploader = lunahub.LunaHubBaseUploader

def convert_aged_ar_to_lunahub_format(df):
    '''
    df is long format with the following columns:
        - ['Name', 'Currency', 'Conversion Factor', 
           'Interval (str)', 'Interval',
           'Value (FCY)', 'Value (LCY)'],
    '''
    
    # Make a copy
    df = df.copy()
    
    # Check that expected columns are present
    expected_columns =  [
        'Name', 'Currency', 'Conversion Factor', 
        'Interval (str)', 'Interval',
        'Value (FCY)', 'Value (LCY)']
    
    cols_not_found = set(expected_columns).difference(df.columns)
    if len(cols_not_found) > 0:
        err = f"Unable to find the following columns: f{cols_not_found}."
        raise Exception (err)
        
    # Convert the interval column
    df["LEFTBINVALUE"] = None
    df["RIGHTBINVALUE"] = None
    
    for i in df.index:
        
        # Get the interval info
        interval = df.at[i, "Interval"]
        l, r, closed = interval.left, interval.right, interval.closed
        
        # If the right is inf, we will set it to 999999
        if np.isinf(r):
            r = 999999
        
        # Update
        df.at[i, "LEFTBINVALUE"] = l
        df.at[i, "RIGHTBINVALUE"] = r
    
    # Convert to int
    df["LEFTBINVALUE"] = df["LEFTBINVALUE"].astype(int)
    df["RIGHTBINVALUE"] = df["RIGHTBINVALUE"].astype(int)
    
    # Add meta info
    df["DATE"]            = None
    df["FY"]              = None
    df["CLIENTNUMBER"]    = None
    df["UPLOADER"]        = None
    df["UPLOADDATETIME"]  = None
    df["COMMENTS"]        = None

    # Map column names
    df = df.rename(columns=AGED_AR_TO_LUNAHUB_MAPPER)
    
    return df
    

class AgedReceivablesReader_Format1:
    
    REQUIRED_HEADERS = ["Name", "Currency", "Conversion Factor", "Total Due"]
    
    def __init__(self, 
                 fp, sheet_name = 0,
                 fy_end_date = None,
                 client_number = None,
                 client_name = None,
                 variance_threshold = 1E-9,
                 insert_to_lunahub = False):
        
        self.fp = fp
        self.sheet_name = sheet_name
        self.fy_end_date = fy_end_date
        self.client_number = client_number
        self.client_name = client_name
        self.variance_threshold = variance_threshold
        self.insert_to_lunahub = insert_to_lunahub
                
    def main(self):
        
        # Read and process data
        self.read_data_from_file()
        self.process_data()
        
        # Load to lunahub if required
        if self.insert_to_lunahub:
            self.load_lunahub()
        
     
    def load_lunahub(self):
        
        # Check that all attributes are provided
        compulsory_attrs = ["fy_end_date", "client_number", "client_name"]
        for attr in compulsory_attrs:
            if getattr(self, attr) is None:
                raise Exception (f"Input {attr} must be specified during initialisation.")
            
        # If okay, then upload
        uploader = AgedReceivablesUploader_To_LunaHub(
            self.df_processed_long_lcy, 
            self.meta_data["Data as at"], # Data date
            self.fy_end_date,
            client_number = self.client_number, 
            client_name = self.client_name)
        
        uploader.main()
        
        self.lunahub_uploader_class = uploader
        
        
    def read_data_from_file(self):
        
        # Read the main df
        df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols(
            self.fp, self.sheet_name)
        
        # Strip empty spaces for strings
        df0 = df0.applymap(lambda s: s.strip() if type(s) is str else s)
        
        self.df0 = df0.copy()
        
        
    def process_data(self):
        
        # Get attr
        df0 = self.df0.copy()
               
        # Get meta data
        meta_data = pd.Series({
            df0.at[r, "A"]: df0.at[r, "B"]
            for r in range(1, 4)
            })
               
        # filter out main data
        df_processed = pyeasylib.pdlib.get_main_table_from_df(
            df0, self.REQUIRED_HEADERS)

        # get the column names
        bin_columns = [
            c for c in df_processed.columns 
            if c not in self.REQUIRED_HEADERS
            ]
        amt_columns = ["Total Due"] + bin_columns
        
        # Save as attr
        self.df_processed           = df_processed.copy()
        self.meta_data              = meta_data.copy()
        self.amt_columns            = amt_columns
        self.bin_columns            = bin_columns
        
        # validate total
        self._validate_total_value()
        
        # process bins
        self._process_bins()
        
        # convert to long format
        self._convert_to_long_format(self.df_processed)
        
        # convert to local currency
        self._convert_to_local_currency()
        
    def _validate_total_value(self):
        
        # Get attr
        df = self.df_processed.copy()
        bin_columns = self.bin_columns
        tot_col = "Total Due"
        
        #
        df["Total Recalculated"] = df[bin_columns].sum(axis=1)
        df["Variance"] = df["Total Recalculated"] - df[tot_col]
        df["Variance (abs)"] = df["Variance"].apply(abs)
        
        # flag out those with variance
        df["Has variance?"] = df["Variance (abs)"] > self.variance_threshold
        
        #
        df_variance = df[df["Has variance?"]]
        
        if df_variance.shape[0] > 0:
            max_variance = df_variance["Variance (abs)"].max()
            error = (
                f"Variance in data - total does not match individual values:\n\n"
                f"{df_variance.T}\n\n"
                f"Please look into the data, or set the variance_threshold to "
                f"> {max_variance:.3f}."
                )
            
            raise Exception (error)
    
    def _convert_to_local_currency(self):
        
        # Get attrs
        df_processed_long = self.df_processed_long.copy()
        
        value_lcy = df_processed_long["Value (FCY)"] * df_processed_long["Conversion Factor"]
        
        df_processed_long["Value (LCY)"] = value_lcy
        
        self.df_processed_long_lcy = df_processed_long.copy()
        
    
    def _process_bins(self):
        
        bin_columns = self.bin_columns
        
        bin_order = pd.Series(range(1, len(bin_columns)+1), index=bin_columns,
                              name = "Order")
        bin_df = bin_order.to_frame()
        
        # create intervals
        bin_df["Interval"] = None
        for bin_str in bin_df.index:
            if "-" in bin_str:
                l, r = bin_str.split("-")
                l = l.strip()
                r = r.strip()
                interval = pd.Interval(int(l), int(r), closed='both')
            elif bin_str.endswith("+"):
                l = bin_str[:-1]
                interval = pd.Interval(int(l), np.inf, closed='left')
            else:
                raise Exception (f"Unexpected bin: {bin_str}.")
            
            bin_df.at[bin_str, "Interval"] = interval
            
        #
        bin_df["lbound"] = bin_df["Interval"].apply(lambda i: i.left)
        bin_df["rbound"] = bin_df["Interval"].apply(lambda i: i.right)
                
        self.bin_df = bin_df.copy()
        
        return bin_df
        
            
    def _convert_to_long_format(self, df):
        
        # get attr
        df = df.copy().reset_index()
        
        # Get the bin cols
        bin_df = self.bin_df
        bin_columns = self.bin_columns
        non_value_columns = [c for c in df.columns if c not in self.amt_columns]
        
        # 
        df_long = df.melt(
            non_value_columns, bin_columns,
            var_name = "Interval (str)",
            value_name = "Value (FCY)")
        
        # Remove those where the value is 0.
        df_long = df_long[df_long["Value (FCY)"] > 0]
        
        # temp order
        df_long["bin_order"] = bin_df["Order"].reindex(
            df_long["Interval (str)"]).values

        # add the bin interval
        df_long["Interval"] = bin_df["Interval"].reindex(
            df_long["Interval (str)"]).values
        
        # sort
        sort_by = ["ExcelRow", "bin_order"]
        df_long = df_long.sort_values(sort_by)
        
        # drop both Excel row and bin order
        for c in sort_by:
            df_long = df_long.drop(c, axis=1)

        # bring value to last row
        new_col_order = [c for c in df_long.columns if c != "Value (FCY)"] + ["Value (FCY)"]
        df_long = df_long[new_col_order]
        
        # reset index, as it is not meaningful anymore
        df_long = df_long.reset_index(drop=True)
        
        self.df_processed_long = df_long.copy()
        

    
    def _split_AR_to_new_groups(self, group_dict):
        '''
        cutoff must lie exactly on one of the interval, as otherwise
        we won't know where to cutoff
        
        for e.g. if there is a $100 at bin 0-30, if the cutoff is 15,
        we won't be able to stratify it.
        
        group_dict = dict of new group name: list of the original bins (str)
        '''
        
        # regroup name
        name = " vs ".join(group_dict.keys())
        
        if not hasattr(self, 'regrouped_dict'):
            
            self.regrouped_dict = {}
        
        if name not in self.regrouped_dict:
                        
            # Get attr
            df_processed_long_lcy = self.df_processed_long_lcy.copy()
            bin_df = self.bin_df.copy()
            
            # Verify that there is no duplicates
            specified_bins = [b for n in group_dict for b in group_dict[n]]
            pyeasylib.assert_no_duplicates(specified_bins)
            
            # Verify all bins are specified
            missing_bins = set(bin_df.index).difference(specified_bins)
            if len(missing_bins) > 0:
                raise Exception (
                    f"Missing bin(s): {list(missing_bins)}.\n\n"
                    f"Please include the bin(s) into the groups correctly: "
                    f"{list(group_dict.keys())}.")
            
            # create a mapper and add a new group column
            mapper_series = pd.Series([n for n in group_dict for b in group_dict[n]],
                                      index = specified_bins)
            df_processed_long_lcy["Group"] = mapper_series.reindex(
                df_processed_long_lcy["Interval (str)"]).values
            
            # save the data
            self.regrouped_dict[name] = df_processed_long_lcy.copy()
            
        return self.regrouped_dict[name]

    def get_AR_by_new_groups(self, group_dict):
        
        df = self._split_AR_to_new_groups(group_dict)
        
        # Value is in foreign currency. only make sense if it's lcy
        return df.pivot_table(
            "Value (LCY)", index="Name", columns="Group", aggfunc="sum")
        
                    
class AgedReceivablesUploader_To_LunaHub(LunaHubBaseUploader):
    
    def __init__(self,
                 df_processed_long_lcy, 
                 date, # Data date
                 fy_end_date,
                 client_number = None, 
                 client_name = None,
                 uploader = None,
                 uploaddatetime = None,
                 lunahub_obj=None):
        
        '''
        df_processed_long_lcy to comprise the following columns:
            - ['Name', 'Currency', 'Conversion Factor', 
               'Interval (str)', 'Interval', 'Value (FCY)', 'Value (LCY)']
        '''
        
        self.df_processed_long_lcy = df_processed_long_lcy.copy()
        self.date                  = date
        self.fy_end_date           = fy_end_date
        self.client_number         = client_number
        self.client_name           = client_name

        # Init parent class        
        LunaHubBaseUploader.__init__(self,
                                     lunahub_obj    = lunahub_obj,
                                     uploader       = uploader,
                                     uploaddatetime = uploaddatetime,
                                     lunahub_config = None)
        

    def main(self):
        
        
        self.upload_client()
        
        self.upload_ar()
        
        
    def upload_ar(self):
        
        # Convert to lunahub format
        df = convert_aged_ar_to_lunahub_format(self.df_processed_long_lcy)
        
        # add meta info
        df["DATE"]              = self.date
        df["FY"]                = self.date.year
        df["CLIENTNUMBER"]      = self.client_number
        df["UPLOADER"]          = self.uploader
        df["UPLOADDATETIME"]    = self.uploaddatetime
        
        # Reorder
        cols = ["NAME", "LEFTBINVALUE", "RIGHTBINVALUE",
                "CURRENCY", "CONVERSIONFACTOR", "VALUEFCY", "VALUELCY",
                "DATE", "FY", "CLIENTNUMBER",
                "UPLOADER", "UPLOADDATETIME", "COMMENTS"]
        df = df[cols]
        
        # Load to lunahub
        self.lunahub_obj.insert_dataframe('ar_aged', df)
        
    def upload_client(self):       
        
        # --------------------------------------------------
        # load client table
        client_uploader = tables.client.ClientInfoUploader_To_LunaHub(
            self.client_number, self.client_name, self.fy_end_date, 
            uploader = self.uploader, uploaddatetime = self.uploaddatetime,
            lunahub_obj = self.lunahub_obj,
            force_insert = False
            )
        client_uploader.main()

        #-------------------------------------------------------
        
class AgedReceivablesLoader_From_LunaHub:
    
    def __init__(self, 
                 client_number,
                 fy,
                 uploaddatetime = None,
                 lunahub_obj = None):
        '''
        specify uploaddatetime (in str) when there are multiple versions of the same data.
        '''
        
        self.client_number  = client_number
        self.fy             = fy
        self.uploaddatetime = uploaddatetime
        self.lunahub_obj    = lunahub_obj
        
        # Initialise lunahub obj if None
        if self.lunahub_obj is None:
            self.lunahub_obj = lunahub.LunaHubConnector(**lunahub.LUNAHUB_CONFIG)
            
        # Main
        self.main()
            
    def main(self):
        
        self.read_data()
        self.process()

    def read_data(self):
        
        df0 = self.lunahub_obj.read_table("ar_aged")
        
        # Check client
        is_client = df0["CLIENTNUMBER"] == self.client_number
        if not is_client.any():
            raise Exception ("Data not found for client_number={self.client_number}.")
        
        # Check FY
        is_fy = df0["FY"] == self.fy
        if not is_fy.any():
            raise Exception ("Data not found for fy={self.fy}.")
    
        # Filter
        df = df0[is_client & is_fy]
        
        # Filter by uploaddatettime
        if self.uploaddatetime is not None:
            
            dt = pd.to_datetime(self.uploaddatetime)
            df = df[df["UPLOADDATETIME"] == dt]
            
            if df.shape[0] == 0:
                raise Exception ("No data found.")        
        
        # Check if there are multiple upload dates
        upload_info = df[["UPLOADDATETIME", "COMMENTS"]].drop_duplicates()
        if upload_info.shape[0] > 1:
            raise Exception (
                f"Multiple uploads for the data:\n\n{upload_info.__repr__()}"
                "\n\nPlease specify the uploaddatetime during initialisation.")
            
        self.df0 = df0
        self.df = df
        
    def process(self):
        # 1) add the interval cols
        # 2) format the cols
        
        # Get
        df = self.df
        
        # Create the interval col
        df["Interval (str)"] = None
        df["Interval"]       = None
        
        for i in df.index:
            
            lval = df.at[i, "LEFTBINVALUE"]
            rval = df.at[i, "RIGHTBINVALUE"]
            
            if rval == 999999:
                intervalstr = f"{lval}+"
                interval = pd.Interval(lval, np.inf, closed='left')
            else:
                intervalstr = f"{lval} - {rval}"
                interval = pd.Interval(lval, rval, closed='both')
            
            # Update
            df.at[i, "Interval (str)"] = intervalstr
            df.at[i, "Interval"] = interval
            
        #
        mapper = {v: k for k, v in AGED_AR_TO_LUNAHUB_MAPPER.items()}
        df = df.rename(columns=mapper)
            
        # reorder cols
        cols = list(mapper.values())
        cols = cols + [c for c in df.columns if c not in cols]
    
        df = df[cols]
        
        self.df = df
        
    
        
if __name__ == "__main__":
    
    if False:
        ar_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\Account receivables listing.xlsx"
        # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\myer gold\Account receivables listing.xlsx
        ar = ARListing(ar_fp, sheet_name="5201 AR")
    
        df3 = ar.process_ar()

    
    # Tester for AgedReceivablesReader_Format1
    if False:
        
        # Specify the file location       
        fp = "../templates/aged_receivables.xlsx"
        sheet_name = "format1"
        
        # Specify the variance threshold - this validates the total column with 
        # the sum of the bins. Try to set to 0 and see what happens.
        variance_threshold = 0.1 #
        
        # Initialise the class
        self = AgedReceivablesReader_Format1(
            fp, sheet_name, 
            fy_end_date = pd.to_datetime("31 Jan 2020"),
            client_number = 1,
            client_name = "ABC PTE LTD",
            variance_threshold=variance_threshold,insert_to_lunahub=True)
        
        # Run main -> this process all the data
        self.main()
        
        # output:
        # df_processed_lcy      -> data in wide format in local currency
        # df_processed_lcy_long -> data in long format in local currency

        # Next, we specify a new grouping (based on 90 split)
        group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                      ">90": ["91 - 120", "121 - 150", "150+"]}

        # Then we get the AR by company (index) and by new bins (columns)
        ar_by_new_grouping = self.get_AR_by_new_groups(group_dict)
        
    
    # Tester to extract
    if True:
        
        client_number = 1
        fy = 2022
        uploaddatetime = '2023-12-08 18:39:03.533'
        lunahub_obj = None
        
        self = AgedReceivablesLoader_From_LunaHub(client_number, fy, 
                                                  uploaddatetime = uploaddatetime,
                                                  lunahub_obj = lunahub_obj)