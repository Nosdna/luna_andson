import pandas as pd

def convert_string_to_interval(s):
    '''
    s = '4900.2' or '2000-3000'
    '''

    if "-" in s:
        l, r = s.split("-")
        l = float(l.strip())
        r = float(r.strip())
    else:
        l = float(s)
        r = l
    interval = pd.Interval(l, r, closed='both')
    
    return interval

def convert_list_of_string_to_interval(string_list):
    
    interval_list = [convert_string_to_interval(s) for s in string_list]
    
    return interval_list


if __name__ == "__main__":
    
    convert_string_to_interval('24.3-  33')