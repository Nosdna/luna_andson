import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment
from openpyxl.formatting.rule import CellIsRule
import pandas as pd
from copy import copy
from datetime import datetime

import luna
from luna.fsvi.mas.template_reader import MASTemplateReader_Form1
from luna.lunahub import tables
import os

class OutputFormatter:

    COLUMN_MAPPER = {"target_ocr_amt_excelcol"      : "F",
                     "target_ocr_subtotal_excelcol" : "G",
                     "target_amt_excelcol"          : "H",
                     "target_subtotal_excelcol"     : "I",
                     "target_var_amt_excelcol"      : "J",
                     "target_var_subtotal_excelcol" : "K",
                     "target_ls_amt_excelcol"       : "L",
                     "target_ls_subtotal_excelcol"  : "M",
                     "target_varname_excelcol"      : "N"
                     }
    
    def __init__(self, input_fp, ocr_fp, output_fp,
                 client_class, aic_name = "", mic_name = ""):

        self.input_fp       = input_fp
        self.ocr_fp         = ocr_fp
        self.output_fp      = output_fp
        self.client_class   = client_class
        self.aic_name       = aic_name
        self.mic_name       = mic_name
        
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
        self.ocr_df = pd.read_excel(self.ocr_fp)

    
    def build_varname_to_values(self, df):
        
        df = df.copy()
    
        varname_to_values = df[
            ["var_name", "Amount", "Subtotal"]
            ].dropna(subset=["var_name"])
        
        # Check that var_name is unique
        if varname_to_values["var_name"].is_unique:
            varname_to_values = varname_to_values.set_index("var_name")
        else:
            raise Exception ("Variable name is not unique.")
                    
        # self.varname_to_values = varname_to_values.copy()
        
        return varname_to_values
    
    def _load_client_info(self):

        self.client_name = self.client_class.retrieve_client_info("CLIENTNAME")

    def _copy_columns(self, ws,
                      src_excelcol, dst_excelcol,
                      src_col_name, dst_col_name,
                      dst_value = None
                      ):
        for src, dst in zip(ws[f"{src_excelcol}:{src_excelcol}"], ws[f"{dst_excelcol}:{dst_excelcol}"]):
            if src.value == src_col_name:
                dst.value = dst_col_name
                dst.font = Font(bold = True)
                src.font = Font(bold = True)
            elif src.value is not None and dst_value is None:
                dst.value = src.value
                dst.border = copy(src.border)
            elif src.value is not None and dst_value is not None:
                dst.value = dst_value
                dst.border = copy(src.border)
            else:
                dst.value = None
                dst.border = copy(src.border)

    def _copy_column_style(self, ws,
                           src_excelcol, dst_excelcol
                           ):
        for src, dst in zip(ws[f"{src_excelcol}:{src_excelcol}"], ws[f"{dst_excelcol}:{dst_excelcol}"]):
            if src.value is not None:
                dst.border = copy(src.border)

    def _create_header(self, ws):

        # Unmerge all cells
        for merge in list(ws.merged_cells):
            ws.unmerge_cells(range_string=str(merge))

        ws.delete_rows(idx = 1, amount = 5)
        ws.insert_rows(idx = 0, amount = 5)

        # for idx in range(4, 6):
        #     ws.row_dimensions[idx].hidden = True #TODO: this should be temporary

        for merge in list(ws.merged_cells):
            ws.unmerge_cells(range_string=str(merge))

        row = 1

        # header title
        ws[f"A{row}"].value = "Form 1 - Statement of assets and liabilities"
        ws[f"A{row}"].font = Font(size = "16", bold = True)
        ws[f"A{row}"].fill = PatternFill("solid", fgColor="00C0C0C0")
        ws[f"A{row}"].border = Border(left   = Side(style = "thin"),
                                 right  = Side(style = "thin"),
                                 top    = Side(style = "thin"),
                                 bottom = Side(style = "thin")
                                 )
        ws.merge_cells(f"A{row}:F{row}")

        date_of_analysis = datetime.now().strftime("%d/%m/%Y")

        dct_of_fields = {"Client name"      : self.client_name,
                         "Date of analysis" : date_of_analysis,
                         "Prepared by"      : self.aic_name,
                         "Reviewed by"      : self.mic_name
                         }
        
        for key in dct_of_fields:
            
            row += 1

            self._create_header_row_field(key, dct_of_fields[key], row, ws)

    def _create_header_row_field(self, field_title, field_value, row, ws):

        border_style = Border(left   = Side(style = "thin"),
                              right  = Side(style = "thin"),
                              top    = Side(style = "thin"),
                              bottom = Side(style = "thin")
                              )
    
        ws[f"A{row}"].value = field_title
        ws[f"A{row}"].fill = PatternFill("solid", fgColor="00C0C0C0")
        ws[f"A{row}"].font = Font(bold = True)
        ws[f"A{row}"].border = border_style
        ws.merge_cells(f"A{row}:C{row}")

        ws[f"D{row}"].value = field_value
        ws[f"D{row}"].fill = PatternFill("solid", fgColor="00FFFFFF")
        ws[f"D{row}"].border = border_style
        ws.merge_cells(f"D{row}:F{row}")
           
    def write_output(self):
        templ_wb = openpyxl.load_workbook(self.template_fp)
        templ_ws = templ_wb[self.template_sheetname]

        self._load_client_info()
        
        sheets_to_remove = [sheet_name for sheet_name in templ_wb.sheetnames if sheet_name != self.template_sheetname]        
        
        for sheet_name in sheets_to_remove:
            templ_wb.remove(templ_wb[sheet_name])
        
        # Template as df
        colname_to_excelcol = self.template_class.colname_to_excelcol
        varname_to_index = self.template_class.varname_to_index
        
        # Get the data
        varname_to_values = self.build_varname_to_values(self.input_df)
        self.varname_to_values = varname_to_values.copy()

        # Get data from ocr
        varname_to_values_ocr = self.build_varname_to_values(self.ocr_df)
        self.varname_to_values_ocr = varname_to_values_ocr.copy()
        
        # Save column index
        amt_excelcol = colname_to_excelcol.at["Amount"]
        subtotal_excelcol = colname_to_excelcol.at["Subtotal"]
        varname_excelcol = colname_to_excelcol.at["var_name"]
        
        # Initialise excelcol
        target_varname_excelcol         = OutputFormatter.COLUMN_MAPPER.get("target_varname_excelcol")
        target_ls_amt_excelcol          = OutputFormatter.COLUMN_MAPPER.get("target_ls_amt_excelcol")
        target_ls_subtotal_excelcol     = OutputFormatter.COLUMN_MAPPER.get("target_ls_subtotal_excelcol")
        target_ocr_amt_excelcol         = OutputFormatter.COLUMN_MAPPER.get("target_ocr_amt_excelcol")
        target_ocr_subtotal_excelcol    = OutputFormatter.COLUMN_MAPPER.get("target_ocr_subtotal_excelcol")
        target_var_amt_excelcol         = OutputFormatter.COLUMN_MAPPER.get("target_var_amt_excelcol")
        target_var_subtotal_excelcol    = OutputFormatter.COLUMN_MAPPER.get("target_var_subtotal_excelcol")
        target_amt_excelcol             = OutputFormatter.COLUMN_MAPPER.get("target_amt_excelcol")
        target_subtotal_excelcol        = OutputFormatter.COLUMN_MAPPER.get("target_subtotal_excelcol")

        # Copy to another column
        self._copy_columns(templ_ws,
                           varname_excelcol, target_varname_excelcol,
                           "var_name", "var_name",
                           dst_value = None)
        self._copy_columns(templ_ws,
                           amt_excelcol, target_amt_excelcol,
                           "Amount", "Amount (Recal)",
                           dst_value = None)
        self._copy_columns(templ_ws,
                           subtotal_excelcol, target_subtotal_excelcol,
                           "Subtotal", "Subtotal (Recal)",
                           dst_value = None)
        self._copy_columns(templ_ws,
                           amt_excelcol, target_ls_amt_excelcol,
                           "Amount", "Amount (L/S)",
                           dst_value = None)
        self._copy_columns(templ_ws,
                           subtotal_excelcol, target_ls_subtotal_excelcol,
                           "Subtotal", "Subtotal (L/S)",
                           dst_value = None)
        self._copy_columns(templ_ws,
                           amt_excelcol, target_ocr_amt_excelcol,
                           "Amount", "Amount (Form 1)",
                           dst_value = "")
        self._copy_columns(templ_ws,
                           subtotal_excelcol, target_ocr_subtotal_excelcol,
                           "Subtotal", "Subtotal (Form 1)",
                           dst_value = "")
        
        amt_excelcol = target_amt_excelcol
        subtotal_excelcol = target_subtotal_excelcol

        # Update amount and subtotal column with recalculated values
        for varname in varname_to_values.index:
            amt = varname_to_values.at[varname, "Amount"]
            subtotal = varname_to_values.at[varname, "Subtotal"]
            
            # Get the location to update
            row = varname_to_index.at[varname]
            
            # Update
            templ_ws[f"{amt_excelcol}{row}"].value = amt
            templ_ws[f"{subtotal_excelcol}{row}"].value = subtotal

            if templ_ws[f"{target_ls_amt_excelcol}{row}"].value in [999999999, '999999999']:
                templ_ws[f"{target_ls_amt_excelcol}{row}"].value = None
            
            if templ_ws[f"{target_ls_subtotal_excelcol}{row}"].value in [999999999, '999999999']:
                templ_ws[f"{target_ls_subtotal_excelcol}{row}"].value = None

            if templ_ws[f"{amt_excelcol}{row}"].value == amt:
                templ_ws[f"{target_var_amt_excelcol}{row}"].value = f"= {target_ocr_amt_excelcol}{row} - {amt_excelcol}{row}"

            if templ_ws[f"{subtotal_excelcol}{row}"].value == subtotal:
                templ_ws[f"{target_var_subtotal_excelcol}{row}"].value = f"= {target_ocr_subtotal_excelcol}{row} - {subtotal_excelcol}{row}"

            # Update format
            templ_ws[f"{amt_excelcol}{row}"].number_format = '#,##0.00'
            templ_ws[f"{subtotal_excelcol}{row}"].number_format = '#,##0.00'
            templ_ws[f"{target_var_amt_excelcol}{row}"].number_format = '#,##0.00'
            templ_ws[f"{target_var_subtotal_excelcol}{row}"].number_format = '#,##0.00'
            templ_ws[f"{target_ocr_amt_excelcol}{row}"].number_format = '#,##0.00'
            templ_ws[f"{target_ocr_subtotal_excelcol}{row}"].number_format = '#,##0.00'

        for varname in varname_to_values_ocr.index:
            amt_ocr = varname_to_values_ocr.at[varname, "Amount"]
            subtotal_ocr = varname_to_values_ocr.at[varname, "Subtotal"]
            
            # Get the location to update
            row = varname_to_index.at[varname]
            
            # Update
            templ_ws[f"{target_ocr_amt_excelcol}{row}"].value = amt_ocr
            templ_ws[f"{target_ocr_subtotal_excelcol}{row}"].value = subtotal_ocr

        self._copy_column_style(templ_ws, amt_excelcol, target_var_amt_excelcol)
        self._copy_column_style(templ_ws, subtotal_excelcol, target_var_subtotal_excelcol)

        # recalculated_excelcol
        recalculated_fill = PatternFill("solid", fgColor="00CCFFFF")
        recalculated_cols = templ_ws[f"{amt_excelcol}7:{subtotal_excelcol}238"] #TODO: currently hardcoded. to update
        recalculated_cols = list(recalculated_cols)
        for cell in recalculated_cols:
            cell[0].fill = recalculated_fill
            cell[1].fill = recalculated_fill

        # ocr_excelcol
        ocr_fill = PatternFill("solid", fgColor="00CCCCFF")
        ocr_cols = templ_ws[f"{target_ocr_amt_excelcol}7:{target_ocr_subtotal_excelcol}238"] #TODO: currently hardcoded. to update
        ocr_cols = list(ocr_cols)
        for cell in ocr_cols:
            cell[0].fill = ocr_fill
            cell[1].fill = ocr_fill

        # var_excelcol
        var_fill = PatternFill("solid", fgColor="00CCFFCC")
        var_cols = templ_ws[f"{target_var_amt_excelcol}7:{target_var_subtotal_excelcol}238"] #TODO: currently hardcoded. to update
        var_cols = list(var_cols)
        for cell in var_cols:
            cell[0].fill = var_fill
            cell[1].fill = var_fill

        # conditional formatting
        redFill = PatternFill(start_color='EE1111',
                              end_color='EE1111',
                              fill_type='solid')
        templ_ws.conditional_formatting.add(f"{target_var_amt_excelcol}8:{target_var_subtotal_excelcol}238",
                                      CellIsRule(operator='greaterThan', formula=['0.01'], stopIfTrue=True, fill=redFill))
        templ_ws.conditional_formatting.add(f"{target_var_amt_excelcol}8:{target_var_subtotal_excelcol}238",
                                      CellIsRule(operator='lessThan', formula=['-0.01'], stopIfTrue=True, fill=redFill))
        templ_ws[f"{target_var_amt_excelcol}7"].value = "Amount (Var)"
        templ_ws[f"{target_var_amt_excelcol}7"].font = Font(bold = True)
        templ_ws[f"{target_var_subtotal_excelcol}7"].value = "Subtotal (Var)"
        templ_ws[f"{target_var_subtotal_excelcol}7"].font = Font(bold = True)

        templ_ws.column_dimensions[target_varname_excelcol].hidden = True
        self._create_header(templ_ws)
        templ_ws.title = "Form 1 (Recalculated)"

        templ_wb.save(self.output_fp)
        templ_wb.close()

if __name__ == "__main__":

    if True: # Testing
        client_no   = 71679
        fy          = 2022

        input_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_71679_2022.xlsx"
        ocr_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_71679_2022_ocr.xlsx"
        output_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_form1_formatted_71679_2022.xlsx"
        
        client_class = tables.client.ClientInfoLoader_From_LunaHub(client_no)

        aic_name = "John Smith"
        mic_name = "Jane Doe"
        
        self = OutputFormatter(input_fp     = input_fp,
                               output_fp    = output_fp,
                               ocr_fp       = ocr_fp,
                               client_class = client_class,
                               aic_name     = aic_name,
                               mic_name     = mic_name
                               )
        
    if True:
        import webbrowser
        webbrowser.open(output_fp)
    
