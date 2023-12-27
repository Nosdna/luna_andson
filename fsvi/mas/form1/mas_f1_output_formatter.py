import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd

class OutputFormatter:

    def __init__(self, input_fp, template_fp, output_fp):

        self.input_fp       = input_fp
        self.template_fp    = template_fp
        self.output_fp      = output_fp
        self.main()

    def main(self):

        self.read_files()
        self.write_output()

    def read_files(self):

        self.input_df = pd.read_excel(self.input_fp)
        self.template_df = pd.read_excel(self.template_fp,
                                         sheet_name = "Form 1 - TB mapping",
                                         skiprows   = 6
                                         )
        
    def write_output(self):
        templ_wb = openpyxl.load_workbook(self.template_fp)
        ws_name = "Form 1 - TB mapping"
        templ_ws = templ_wb[ws_name]

        sheets_to_remove = [sheet_name for sheet_name in templ_wb.sheetnames if sheet_name != ws_name]
        for sheet_name in sheets_to_remove:
            templ_wb.remove(templ_wb[sheet_name])

        for index, row in self.input_df.iterrows():
            var_name = row["var_name"]
            amount = row["Amount"]
            subtotal = row["Subtotal"]

            for excel_row in range(1, templ_ws.max_row + 1):
                if templ_ws.cell(row=excel_row, column=8).value == var_name:
                    templ_ws.cell(row=excel_row, column=6, value=amount)
                    templ_ws.cell(row=excel_row, column=7, value=subtotal)
                    break 

        templ_ws.column_dimensions['H'].hidden = True
    
        

        templ_ws.title = "Form 1 (Recalculated)"

        templ_wb.save(self.output_fp)
        templ_wb.close()

if __name__ == "__main__":

    if False: # Testing
        client_no   = 71679
        fy          = 2022

        input_fp    = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_7167_2022.xlsx"
        template_fp = r"D:\workspace\luna\parameters\mas_forms_tb_mapping.xlsx"
        output_fp   = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_formatted_71679_2022.xlsx"

        self = OutputFormatter(input_fp     = input_fp,
                               template_fp  = template_fp,
                               output_fp    = output_fp
                               )
        
