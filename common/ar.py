import pandas as pd
import numpy as np
import pyeasylib
import re

class AgedReceivablesReader_Format1:
    
    REQUIRED_HEADERS = ["Name", "Currency", "Conversion Factor", "Total Due"]
    
    def __init__(self, fp, sheet_name = 0, variance_threshold = 1E-9):
        
        self.fp = fp
        self.sheet_name = sheet_name
        self.variance_threshold = variance_threshold
        
        self.main()
        
    def main(self):
        
        self.read_data_from_file()
        
        self.process_data()
        
        
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
                    f"Missing bin(s): {list(missing_bins)}. "
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
        
                    

class ARListing_DACIA:
    
    def __init__(self, file, sheet_name):
        
        self.file = file
        self.sheet_name = sheet_name
            
    def process_ar(self):
        self.df3 = pd.read_excel(self.file, sheet_name = self.sheet_name)

        for i in range(35,37):
            self.df3.at[i,"Unnamed: 3"] = self.df3.at[i, "Unnamed: 2"]

        self.df3 = self.df3.iloc[8:37, [1,3]]
        self.df3 = self.df3.dropna()
        self.df3.columns = self.df3.iloc[0]
        self.df3 = self.df3.drop(self.df3.index[0])
        return self.df3
    
    def process_ageing(self):
        self.ar_ageing = pd.read_excel(self.file, sheet_name = self.sheet_name)

        self.ar_ageing = self.ar_ageing.iloc[8:37, 1:9]
        self.ar_ageing.columns = self.ar_ageing.iloc[0]
        self.ar_ageing.drop(self.ar_ageing.index[0], inplace = True)

        pd.set_option('display.float_format', '{:.4f}'.format)

        self.ar_ageing.iloc[:23,1:] = self.ar_ageing.iloc[:23,1:].apply(self.convert)

        self.ar_ageing.drop([30,31,32,33,36], inplace = True)

        return self.ar_ageing
    
    def convert(self, x):
        pattern = "converted.*[at|@].*(\d{1}\.\d{4})"
        rate = float(re.findall(pattern, str(self.ar_ageing["Name"]))[0])
        x = pd.to_numeric(x, errors = 'coerce') * rate
        return x
    
if __name__ == "__main__":
    
    if False:
        ar_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\Account receivables listing.xlsx"
        # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\myer gold\Account receivables listing.xlsx
        ar = ARListing(ar_fp, sheet_name="5201 AR")
    
        df3 = ar.process_ar()

    
    # Tester for AgedReceivablesReader_Format1
    if True:
        
        # Specify the file location
        fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\aged_receivables_template.xlsx"
        
        fp = "../templates/aged_receivables.xlsx"
        sheet_name = "format1"
        
        # Specify the variance threshold - this validates the total column with 
        # the sum of the bins. Try to set to 0 and see what happens.
        variance_threshold = 0.1 #
        
        # Initialise the class
        self = AgedReceivablesReader_Format1(
            fp, sheet_name, variance_threshold=variance_threshold)
        
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