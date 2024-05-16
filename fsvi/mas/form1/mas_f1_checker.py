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

class MASForm1_Validator:

    def __init__(self,
                 gold_fp,
                 input_fp):

        self.gold_fp    = gold_fp
        self.input_fp   = input_fp

        self.main()

    def main(self):

        self.read_raw_f1()
        self.process_raw_f1()
        self.compare_f1()
        self.print_results()

    def read_raw_f1(self):

        self.gold_raw_f1 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.gold_fp)
        self.input_raw_f1 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.input_fp)

    def process_raw_f1(self):

        gold_f1, gold_headers = self._rename_raw_f1(self.gold_raw_f1)
        input_f1, input_headers = self._rename_raw_f1(self.input_raw_f1)

        gold_filtered_f1 = self._filter_f1(gold_f1, gold_headers)
        input_filtered_f1 = self._filter_f1(input_f1, input_headers)

        self.gold_filtered_f1 = gold_filtered_f1
        self.input_filtered_f1 = input_filtered_f1

    def _rename_raw_f1(self, df):
        
        col_header_row = df.filter(items=[8,9], axis = 0)
        col_header_row.ffill(axis = 1, inplace = True)
        col_header_row = col_header_row.groupby(by = ['A'], dropna = False).agg({'F':lambda x: ' - '.join(x.unique()),
                                                                                 'G':lambda x: ' - '.join(x.unique()),
                                                                                 'H':lambda x: ' - '.join(x.unique()),
                                                                                 'I':lambda x: ' - '.join(x.unique()),
                                                                                 'J':lambda x: ' - '.join(x.unique()),
                                                                                 'K':lambda x: ' - '.join(x.unique()),
                                                                                 'N':'last'
                                                                                 })
        col_headers = col_header_row.reset_index().T[0].dropna().to_dict()

        renamed_df = df.rename(col_headers, axis = 1)

        return renamed_df, col_headers

    def _filter_f1(self, df, headers):

        filtered_df = df[~df['var_name'].isna()]

        filtered_df = filtered_df[list(headers.values())].set_index('var_name')

        filtered_df = filtered_df[~(filtered_df.index=='var_name')]

        return filtered_df

    def compare_f1(self):

        gold_f1 = self.gold_filtered_f1.copy()
        input_f1 = self.input_filtered_f1.copy()
        merged_f1 = pd.merge(gold_f1, input_f1,
                             how        = "outer",
                             on         = 'var_name',
                             suffixes   = ["_gold", "_input"]
                             )
        
        merged_f1.fillna(0, inplace = True)
        base_cols = set([col.split('_')[0] for col in merged_f1.columns])
        for base_col in base_cols:
            if f'{base_col}_gold' in merged_f1.columns and f'{base_col}_input' in merged_f1.columns:
                new_col = f'{base_col}_check'
                merged_f1[new_col] = merged_f1[f'{base_col}_gold'] - merged_f1[f'{base_col}_input'] 

        check_cols = [col for col in merged_f1.columns if '_check' in col]
        merged_f1 = merged_f1[check_cols]

        self.merged_f1_detailed = merged_f1
        self.merged_f1_summary = merged_f1.agg('sum')

    def print_results(self):

        summary = self.merged_f1_summary
        summary_check = summary[summary != 0]

        if len(summary_check) > 0:
            print("There are differences between the summary of the gold standard and the input file for Form 1. \n"
                  f"Please check the following columns: {summary_check}")
            
        else:
            print("There are no differences between the summary of the gold standard and the input file for Form 1.")

        detail = self.merged_f1_detailed
        detail_check = detail[detail != 0].dropna(how = 'all')

        if len(detail_check) > 0:
            print("There are differences between the detail of the gold standard and the input file for Form 1. \n"
                  f"Please check the following columns: {detail_check}")
            
        else:
            print("There are no differences between the detail of the gold standard and the input file for Form 1.")
    


if __name__ == '__main__':

    client_code = 7167
    fy = 2022

    gold_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form1_formatted_{client_code}_{fy}.xlsx"
    input_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form1_formatted_{client_code}_{fy}.xlsx"

    self = MASForm1_Validator(gold_fp,
                              input_fp)
