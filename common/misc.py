import pandas as pd
import inspect
import numpy as np

def convert_string_to_interval(s):
    '''
    s = '4900.2' or '2000-3000'
    '''
    try:
        if "-" in s:
            l, r = s.split("-")
            l = float(l.strip())
            r = float(r.strip())
        else:
            l = float(s)
            r = l
        interval = pd.Interval(l, r, closed='both')
    except Exception as e:
        raise Exception (f"Unable to convert to interval for: {s}.\n{str(e)}")
    
    return interval

def convert_list_of_string_to_interval(string_list):
    
    interval_list = [convert_string_to_interval(s) for s in string_list]
    
    return interval_list


def convert_binstrs_to_bin_df(binstr_list):
    '''
    E.g. 
    binstr_list = ['0 - 30', '31 - 60', '61 - 90', '91 - 120', '121 - 150', '150+']
    '''
    
    # Take the unique
    bins = list(set(binstr_list))
    
    # Set an arbitrary order first    
    bin_order = pd.Series(range(1, len(bins)+1), index=bins,
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
    
    # Order and reset the bin order
    bin_df = bin_df.sort_values("lbound")
    bin_df["Order"] = range(1, bin_df.shape[0]+1)
            
    return bin_df


def get_my_name():
    '''
    Returns the name of the method when this method is called.
    '''
        
    #https://stackoverflow.com/questions/5067604/determine-function-name-from-within-that-function-without-using-traceback
    
    return inspect.currentframe().f_back.f_code.co_name



if __name__ == "__main__":
    
    convert_string_to_interval('24.3-  33')
