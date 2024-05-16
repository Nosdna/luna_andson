# Import standard libs
import os
import datetime
import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
import sys


# Import luna package, fsvi package and pyeasylib
import pyeasylib
import luna
import luna.common as common
import luna.fsvi as fsvi
import luna.lunahub as lunahub

# TODO: to add validation checks on length of each dataframe (gold, input)

class MASForm3_Validator:

    def __init__(self,
                 gold_fp,
                 input_fp):

        self.gold_fp    = gold_fp
        self.input_fp   = input_fp

        self.main()

    def main(self):

        self.read_raw_f3()
        self.process_raw_f3()
        self.compare_f3()
        self.print_results()

    def read_raw_f3(self):

        self.gold_raw_f3 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.gold_fp)
        self.input_raw_f3 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.input_fp)

    def process_raw_f3(self):

        gold_f3, gold_headers = self._rename_raw_f3(self.gold_raw_f3)
        input_f3, input_headers = self._rename_raw_f3(self.input_raw_f3)

        gold_filtered_f3 = self._filter_f3(gold_f3, gold_headers)
        input_filtered_f3 = self._filter_f3(input_f3, input_headers)

        self.gold_filtered_f3 = gold_filtered_f3
        self.input_filtered_f3 = input_filtered_f3

    def _rename_raw_f3(self, df):

        col_header_row = df.filter(items=[8,9], axis = 0)
        col_header_row.ffill(axis = 1, inplace = True)
        col_header_row = col_header_row.groupby(by = ['A'], dropna = False).agg({'E':lambda x: ' - '.join(x.unique()),
                                                                                 'F':lambda x: ' - '.join(x.unique()),
                                                                                 'G':lambda x: ' - '.join(x.unique()),
                                                                                 'H':lambda x: ' - '.join(x.unique()),
                                                                                 'I':lambda x: ' - '.join(x.unique()),
                                                                                 'J':lambda x: ' - '.join(x.unique()),
                                                                                #  'K':lambda x: ' - '.join(x.unique()),
                                                                                #  'L':lambda x: ' - '.join(x.unique()),
                                                                                 'M':'last'
                                                                                 })
        col_headers = col_header_row.reset_index().T[0].dropna().to_dict()

        renamed_df = df.rename(col_headers, axis = 1)

        return renamed_df, col_headers

    def _filter_f3(self, df, headers):

        filtered_df = df[~df['var_name'].isna()]

        filtered_df = filtered_df[list(headers.values())].set_index('var_name')

        filtered_df = filtered_df[~(filtered_df.index=='var_name')]

        return filtered_df

    def compare_f3(self):

        gold_f3 = self.gold_filtered_f3.copy()
        input_f3 = self.input_filtered_f3.copy()
        merged_f3 = pd.merge(gold_f3, input_f3,
                             how        = "outer",
                             on         = 'var_name',
                             suffixes   = ["_gold", "_input"]
                             )
        merged_f3.fillna(0, inplace = True)

        base_cols = set([col.split('_')[0] for col in merged_f3.columns])
        for base_col in base_cols:
            if f'{base_col}_gold' in merged_f3.columns and f'{base_col}_input' in merged_f3.columns:
                new_col = f'{base_col}_check'
                merged_f3[new_col] = merged_f3[f'{base_col}_gold'] - merged_f3[f'{base_col}_input'] 

        check_cols = [col for col in merged_f3.columns if '_check' in col]
        merged_f3 = merged_f3[check_cols]

        self.merged_f3_detailed = merged_f3
        self.merged_f3_summary = merged_f3.agg('sum')

    def print_results(self):

        summary = self.merged_f3_summary
        summary_check = summary[summary != 0]

        if len(summary_check) > 0:
            print("There are differences between the summary of the gold standard and the input file for Form 3. \n"
                  f"Please check the following columns: {summary_check}")
            
        else:
            print("There are no differences between the summary of the gold standard and the input file for Form 3.")

        detail = self.merged_f3_detailed
        detail_check = detail[detail != 0].dropna(how = 'all')

        if len(detail_check) > 0:
            print("There are differences between the detail of the gold standard and the input file for Form 3. \n"
                  f"Please check the following columns: {detail_check}")
            
        else:
            print("There are no differences between the detail of the gold standard and the input file for Form 3.")
    

if __name__ == '__main__':

    client_code = 7167
    fy = 2022

    gold_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form3_formatted_{client_code}_{fy}.xlsx"
    input_fp = r"D:\workspace\luna\personal_workspace\tmp" + f"\mas_form3_formatted_{client_code}_{fy}.xlsx"

    self = MASForm3_Validator(gold_fp,
                              input_fp)
