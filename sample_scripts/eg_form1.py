'''
Sample script for Dacia and Jia Wey
'''

# Import standard libs
import os
import datetime

# Import luna package and fsvi package

import luna
import luna.common as common
import luna.fsvi as fsvi

class MASForm1_Generator:
    
    def __init__(self, tb_classs, aged_ar_class, mapper_class):
        
        
        self.tb_class       = tb_class
        self.aged_ar_class  = aged_ar_class
        self.mapper_class   = mapper_class
        
    def do_something(self):
        
        pass
    




if __name__ == "__main__":
    
   
        
    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    print (f"Your luna library is at {luna_folderpath}.")
    
    # Get the template folderpath
    template_folderpath = os.path.join(luna_folderpath, "templates")
    
    # AGED RECEIVABLES
    if True:
        aged_receivables_fp = os.path.join(template_folderpath, "aged_receivables.xlsx")
        print (f"Your aged_receivables_fp is at {aged_receivables_fp}.")
        
        # Load the AR class
        aged_ar_class = common.AgedReceivablesReader_Format1(aged_receivables_fp, 
                                                        sheet_name = 0,            # Set the sheet name
                                                        variance_threshold = 1E-9) # To relax criteria if required.
        
        aged_group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                           ">90": ["91 - 120", "121 - 150", "150+"]}
        
        # Then we get the AR by company (index) and by new bins (columns)
        aged_df_by_company = aged_ar_class.get_AR_by_new_groups(aged_group_dict)
        
    # TB
    if True:
        tb_fp = os.path.join(template_folderpath, "tb.xlsx")
        print (f"Your tb_filepath is at {tb_fp}.")
        
        # Load the tb
        fy_end_date = datetime.date(2022, 12, 31)
        tb_class = common.TBReader_ExcelFormat1(tb_fp, 
                                                sheet_name = 0,
                                                fy_end_date = fy_end_date)
        
        
        # Get data by fy
        fy = 2022
        tb2022 = tb_class.get_data_by_fy(fy)
        
    # Form 1 mapping
    if True:
        
        mas_tb_mapping_fp = os.path.join(luna_folderpath, "parameters", "mas_forms_tb_mapping.xlsx")
        print (f"Your mas_tb_mapping_fp is at {mas_tb_mapping_fp}.")
        
        # Load the class
        mapper_class = fsvi.mas.MASTemplateReader_Form1(mas_tb_mapping_fp, sheet_name = "Form 1 - TB mapping")
    
        # process df is here:
        df_processed = mapper_class.df_processed  # need to build methods


    # CLASS
    self = MASForm1_Generator(tb_class, aged_ar_class, mapper_class)