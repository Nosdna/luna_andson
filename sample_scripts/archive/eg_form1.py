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
    
    def __init__(self, 
                 tb_classs, aged_ar_class, mapper_class,
                 fy = 2022):
        
        
        self.tb_class       = tb_class
        self.aged_ar_class  = aged_ar_class
        self.mapper_class   = mapper_class
        self.fy = fy
        
        self.main()
        
    def main(self):
        
        self._map_varname_to_lscodes()

    def _map_varname_to_lscodes(self):
        
        mapper_class = self.mapper_class
        tb_class     = self.tb_class
        
        # get varname to ls code from mapper
        varname_to_lscodes = mapper_class.varname_to_lscodes
        
        # get the tb for the current fy
        tb_df = tb_class.get_data_by_fy(self.fy).copy()
        tb_columns = tb_df.columns
        
        # screen thru
        for varname in varname_to_lscodes.index:
            #varname = "puc_pref_share_noncumulative"
            lscode_intervals = varname_to_lscodes.at[varname]
            
            # Get the true and false matches
            is_overlap, tb_df_true, tb_df_false = tb_class.filter_tb_by_fy_and_ls_codes(
                self.fy, lscode_intervals)
            
            # Update the main table
            tb_df[varname] = is_overlap

        #
        self.tb_columns_main = tb_columns
        self.tb_with_varname = tb_df.copy()

    def filter_tb_by_varname(self, varname):
        
        tb = self.tb_with_varname.copy()
        
        filtered_tb = tb[tb[varname]][self.tb_columns_main]
        
        return filtered_tb
        
        
        


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
        aged_receivables_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\aged_receivables_template.xlsx"
        print (f"Your aged_receivables_fp is at {aged_receivables_fp}.")
        
        # Load the AR class
        aged_ar_class = common.AgedReceivablesReader_Format1(aged_receivables_fp, 
                                                        sheet_name = 0,            # Set the sheet name
                                                        variance_threshold = 0.1) # 1E-9) # To relax criteria if required.
        
        aged_group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                           ">90": ["91 - 120", "121 - 150", "150+"]}
        
        # Then we get the AR by company (index) and by new bins (columns)
        aged_df_by_company = aged_ar_class.get_AR_by_new_groups(aged_group_dict)
        
    # TB
    if True:
        tb_fp = os.path.join(template_folderpath, "tb.xlsx")
        tb_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\Myer Gold Investment Management - 2022 TB.xlsx"
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
    fy=2022
    self = MASForm1_Generator(tb_class, aged_ar_class,
                              mapper_class, fy=fy)
    
    # Get df by varname
    filtered_tb = self.filter_tb_by_varname('current_asset_trade_debt_other')
    
    
    