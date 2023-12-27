import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd

import luna
from luna.fsvi.mas.template_reader import MASTemplateReader_Form1
import os

class OutputFormatter:

    def __init__(self, input_fp, output_fp):

        self.input_fp       = input_fp
        self.output_fp      = output_fp
        
        # Get the template fp
        self.template_fp = os.path.join(
            luna.settings.LUNA_FOLDERPATH,
            "parameters", 
            "mas_forms_tb_mapping.xlsx")
        
        self.template_sheetname = "Form 1 - TB mapping"
        
        self.template_class = MASTemplateReader_Form1(
            self.template_fp, 
            self.template_sheetname)
        
        
        self.main()

    def main(self):

        self.read_files()
        self.write_output()

    def read_files(self):
        # Read the file with the values
        self.input_df = pd.read_excel(self.input_fp)

    
    def build_varname_to_values(self):
        
        #
        input_df = self.input_df.copy()
    
        varname_to_values = input_df[
            ["var_name", "Amount", "Subtotal"]
            ].dropna(subset=["var_name"])
        
        # Check that var_name is unique
        if varname_to_values["var_name"].is_unique:
            varname_to_values = varname_to_values.set_index("var_name")
        else:
            raise Exception ("Variable name is not unique.")
                    
        self.varname_to_values = varname_to_values.copy()
        
        return self.varname_to_values
                
            
    def write_output(self):
        templ_wb = openpyxl.load_workbook(self.template_fp)
        templ_ws = templ_wb[self.template_sheetname]

        sheets_to_remove = [sheet_name for sheet_name in templ_wb.sheetnames if sheet_name != self.template_sheetname]        
        
        #This part causes error 
        for sheet_name in sheets_to_remove:
        #    templ_wb.remove(templ_wb[sheet_name])
            #del templ_wb[sheet_name]
            templ_wb.remove(templ_wb[sheet_name])
        
        # Template as df
        colname_to_excelcol = self.template_class.colname_to_excelcol
        varname_to_index = self.template_class.varname_to_index
        
        # Get the data
        varname_to_values = self.build_varname_to_values()
        
        #
        if True:
            amt_excelcol = colname_to_excelcol.at["Amount"]
            subtotal_excelcol = colname_to_excelcol.at["Subtotal"]
            for varname in varname_to_values.index:
                amt = varname_to_values.at[varname, "Amount"]
                subtotal = varname_to_values.at[varname, "Subtotal"]
                
                # Get the location to update
                row = varname_to_index.at[varname]
                
                # Update
                templ_ws[f"{amt_excelcol}{row}"].value = amt
                templ_ws[f"{subtotal_excelcol}{row}"].value = subtotal
                
                
            templ_ws.column_dimensions['H'].hidden = True
        
            
    
            templ_ws.title = "Form 1 (Recalculated)"

        templ_wb.save(self.output_fp)
        templ_wb.close()

if __name__ == "__main__":

    if True: # Testing
        client_no   = 71679
        fy          = 2022

        #input_fp    = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_7167_2022.xlsx"
        #template_fp = r"D:\workspace\luna\parameters\mas_forms_tb_mapping.xlsx"
        #output_fp   = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_formatted_71679_2022.xlsx"

        input_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\tmp\mas_form1_71679_2022.xlsx"
        output_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\tmp\mas_form1_71679_2022_out.xlsx"
        self = OutputFormatter(input_fp     = input_fp,
                               output_fp    = output_fp
                               )
        
