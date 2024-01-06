# Import standard libs
import os
import datetime
import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
import sys
import logging

# Initialise logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())

# Import luna package and fsvi package

import luna
import luna.common as common


from luna.common.workbk import CreditDropdownList

class MASForm2_Generator:
    
    def __init__(self, tb_class,
                 credit_quality_output_fp,
                 fy = 2022
                 ):
        
        
        self.tb_class                   = tb_class
        self.credit_quality_output_fp   = credit_quality_output_fp
        self.fy                         = fy

        self.main()
        
    def main(self):

        self.create_credit_quality_output()


#### (IV) Adjusted Assets 

    def create_credit_quality_output(self):
        
        tb_df = self.tb_class.get_data_by_fy(self.fy).copy()
        
        interval_list = [pd.Interval(5000,5000,closed = "both")]
        cash = tb_df[tb_df["L/S (interval)"].apply(lambda x:x in interval_list)]

        depo = cash[~cash["Name"].str.contains("(?i)cash")]
        
        new_cols = depo.columns.to_list() + ["Credit Quality Grade 1?"]
        
        depo = depo.reindex(columns=new_cols)

        print("Please tag if the following deposits are credit quality grade 1")
        
        depo.to_excel(self.credit_quality_output_fp, index = False)
        dropdownlist = CreditDropdownList(self.credit_quality_output_fp)
        depo = dropdownlist.create_dropdown_list()
        
        
if __name__ == "__main__":

    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    print (f"Your luna library is at {luna_folderpath}.")
    
    # Get the template folderpath
    template_folderpath = os.path.join(luna_folderpath, "templates")

    # TESTER
    if True:
        client_number   = 7167
        fy              = 2022
       
    ### TB ###
    # Load from file
    if False:
        #tb_fp = os.path.join(template_folderpath, "tb.xlsx")
        tb_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\f2\f2_tb_used.xlsx"

        print (f"Your tb_filepath is at {tb_fp}.")
        
        # Load the tb
        fy_end_date = datetime.date(2022, 12, 31)
        tb_class = common.TBReader_ExcelFormat1(tb_fp, 
                                                sheet_name = 0,
                                                fy_end_date = fy_end_date)
        # Get data by fy
        fy = 2022
        tb2022 = tb_class.get_data_by_fy(fy)
    # Load from LunaHub
    if True:
        tb_class = common.TBLoader_From_LunaHub(client_number, fy)
        
    # CLASS
        
    # Credit quality output fp
    credit_quality_output_fp = rf"D:\workspace\luna\personal_workspace\tmp\mas_f2_{client_number}_{fy}_credit_quality.xlsx"
    
    self = MASForm2_Generator(tb_class,
                              credit_quality_output_fp,
                              fy=fy
                              )
    
    # Get df by varname
    # filtered_tb = self.filter_tb_by_varname('current_asset_trade_debt_other')
    
    # Output to excel 
    # self.outputdf.to_excel("draftf2.xlsx") 


    