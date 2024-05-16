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

# from form2_part2 import MASForm2_Generator_Part2

class MASForm2_Validator:

    def __init__(self,
                 gold_fp,
                 input_fp):

        self.gold_fp    = gold_fp
        self.input_fp   = input_fp

        self.main()

    def main(self):

        self.read_raw_f2()
        self.process_raw_f2()
        self.compare_f2()
        self.print_results()

    def read_raw_f2(self):

        self.gold_raw_f2 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.gold_fp)
        self.input_raw_f2 = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp = self.input_fp)

    def process_raw_f2(self):

        gold_f2, gold_headers = self._rename_raw_f2(self.gold_raw_f2)
        input_f2, input_headers = self._rename_raw_f2(self.input_raw_f2)

        gold_filtered_f2 = self._filter_f2(gold_f2, gold_headers)
        input_filtered_f2 = self._filter_f2(input_f2, input_headers)

        self.gold_filtered_f2 = gold_filtered_f2
        self.input_filtered_f2 = input_filtered_f2

    def _rename_raw_f2(self, df):

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

    def _filter_f2(self, df, headers):

        filtered_df = df[~df['var_name'].isna()]

        filtered_df = filtered_df[list(headers.values())].set_index('var_name')

        filtered_df = filtered_df[~(filtered_df.index=='var_name')]

        return filtered_df

    def compare_f2(self):

        gold_f2 = self.gold_filtered_f2.copy()
        input_f2 = self.input_filtered_f2.copy()
        merged_f2 = pd.merge(gold_f2, input_f2,
                             how        = "outer",
                             on         = 'var_name',
                             suffixes   = ["_gold", "_input"]
                             )
        merged_f2.fillna(0, inplace = True)

        base_cols = set([col.split('_')[0] for col in merged_f2.columns])
        for base_col in base_cols:
            if f'{base_col}_gold' in merged_f2.columns and f'{base_col}_input' in merged_f2.columns:
                new_col = f'{base_col}_check'
                merged_f2[new_col] = merged_f2[f'{base_col}_gold'] - merged_f2[f'{base_col}_input'] 

        check_cols = [col for col in merged_f2.columns if '_check' in col]
        merged_f2 = merged_f2[check_cols]

        self.merged_f2_detailed = merged_f2
        self.merged_f2_summary = merged_f2.agg('sum')

    def print_results(self):

        summary = self.merged_f2_summary
        summary_check = summary[summary != 0]

        if len(summary_check) > 0:
            print("There are differences between the summary of the gold standard and the input file for Form 2. \n"
                  f"Please check the following columns: {summary_check}")
            
        else:
            print("There are no differences between the summary of the gold standard and the input file for Form 2.")

        detail = self.merged_f2_detailed
        detail_check = detail[detail != 0].dropna()

        if len(detail_check) > 0:
            print("There are differences between the detail of the gold standard and the input file for Form 2. \n"
                  f"Please check the following columns: {detail_check}")
            
        else:
            print("There are no differences between the detail of the gold standard and the input file for Form 2.")
    
class MASForm2_AWP_Validator:

    def __init__(self,
                 gold_fp,
                 input_fp):

        self.gold_fp    = gold_fp
        self.input_fp   = input_fp

        self.main()

    def main(self):

        self.read_raw_awp()
        self.process_awp()
        self.compare_trr()
        self.compare_aaa()
        self.print_results()

    def read_raw_awp(self):

        self.gold_raw_awp_aaa = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp          = self.gold_fp,
                                                                                sheet_name  = 'aaa'
                                                                                )
        self.input_raw_awp_aaa = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp         = self.input_fp,
                                                                                 sheet_name = 'aaa'
                                                                                 )
        self.gold_raw_awp_trr = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp         = self.gold_fp,
                                                                                 sheet_name = 'trr'
                                                                                 )
        self.input_raw_awp_trr = pyeasylib.excellib.read_excel_with_xl_rows_cols(fp         = self.input_fp,
                                                                                 sheet_name = 'trr'
                                                                                 )

    def process_awp(self):

        self.gold_awp_trr_agi, self.gold_awp_trr_snf = self._process_raw_awp_trr(self.gold_raw_awp_trr)
        self.input_awp_trr_agi, self.input_awp_trr_snf = self._process_raw_awp_trr(self.input_raw_awp_trr)

        self.gold_awp_aaa = self._process_raw_awp_aaa(self.gold_raw_awp_aaa)
        self.input_awp_aaa = self._process_raw_awp_aaa(self.input_raw_awp_aaa)

    def _process_raw_awp_trr(self, df):

        # annual gross income 3 year table
        start = df[df.iloc[:,0].str.contains("Definition",
                                             na=False)].index[0]
        end = df[df.iloc[:,1].str.contains("Adjusted annual gross income",
                                           na = False)].index[0]
        agi_df = df.iloc[start-1:end+1, 1:5]
        agi_df.columns = agi_df.iloc[0].fillna('field')
        agi_df = agi_df.iloc[2:]
        agi_df.dropna(how = 'all', axis = 1, inplace = True)
        agi_df.dropna(how = 'all', axis = 0, inplace = True)
        agi_df.reset_index(drop = True, inplace = True)

        # securities and futures
        start = df[df.iloc[:,0].str.contains('Paid-up capital',
                                             na=False)].index[0]
        end = df[df.iloc[:,0].str.contains("Definition",
                                           na = False)].index[0]
        snf_df = df.iloc[start-1:end-2, 1:9]
        # snf_df.columns = {'B' : 'field', 'F' : ''}
        # snf_df = snf_df.iloc[1:]
        snf_df.dropna(subset=['B'], inplace = True)
        snf_df.dropna(how = 'all', axis = 1, inplace = True)
        snf_df.dropna(how = 'all', axis = 0, inplace = True)
        snf_df.reset_index(drop = True, inplace = True)

        return agi_df, snf_df
    
    def _process_raw_awp_aaa(self, df):

        start = df[df.iloc[:,0].str.contains("Para 3.3.5 Average Adjusted Assets calculated at the end of each month in a given quarter",
                                             na=False)].index[0]
        end = df[df.iloc[:,1].str.contains("fee receivables owed by a customer account",
                                           na = False)].index[0]
        aaa_df = df.iloc[start:end, 1:6]
        aaa_df.dropna(how = 'all', axis = 1, inplace = True)
        aaa_df.dropna(how = 'all', axis = 0, inplace = True)
        aaa_df.columns = aaa_df.iloc[0].fillna('field')
        aaa_df = aaa_df.iloc[3:]
        aaa_df.reset_index(drop = True, inplace = True)

        return aaa_df
    
    def compare_trr(self):
        
        gold_agi = self.gold_awp_trr_agi.copy()
        input_agi = self.input_awp_trr_agi.copy()
        merged_agi = self._compare_df(gold_agi, input_agi, ['field'])

        gold_snf = self.gold_awp_trr_snf.copy()[['B', 'I']]
        input_snf = self.input_awp_trr_snf.copy()[['B', 'I']]
        merged_snf = self._compare_df(gold_snf, input_snf, ['B'])

        self.merged_agi_detailed = merged_agi
        self.merged_agi_summary = merged_agi.agg('sum')

        self.merged_snf_detailed = merged_snf
        self.merged_snf_summary = merged_snf.agg('sum')

    def compare_aaa(self):

        gold_aaa = self.gold_awp_aaa.copy()
        input_aaa = self.input_awp_aaa.copy()
        merged_aaa = self._compare_df(gold_aaa, input_aaa, ['field'])

        self.merged_aaa_detailed = merged_aaa
        self.merged_aaa_summary = merged_aaa.agg('sum')

    def _compare_df(self, df0, df1, key_lst):

        merged = pd.merge(df0, df1,
                              how = 'outer',
                              on = key_lst,
                              suffixes = ["_gold", "_input"])
        merged.fillna(0, inplace = True)
        base_cols = set([col.split('_')[0] for col in merged.columns])
        for base_col in base_cols:
            if f'{base_col}_gold' in merged.columns and f'{base_col}_input' in merged.columns:
                new_col = f'{base_col}_check'
                merged[new_col] = merged[f'{base_col}_gold'] - merged[f'{base_col}_input'] 

        check_cols = [col for col in merged.columns if '_check' in col]
        merged = merged.set_index(key_lst)[check_cols]

        return merged

    def print_results(self):

        # agi
        self._prepare_results('Annual gross income section of TRR',
                              self.merged_agi_detailed,
                              self.merged_snf_summary
                              )
        
        # snf
        self._prepare_results('General section of TRR',
                              self.merged_snf_detailed,
                              self.merged_snf_summary
                              )
        
        # aaa
        self._prepare_results('AAA tab',
                              self.merged_aaa_detailed,
                              self.merged_aaa_summary
                              )

    def _prepare_results(self, name, detail, summary):

        summary_check = summary[summary != 0]

        if len(summary_check) > 0:
            print(f"There are differences between the summary of the gold standard and the input file for '{name}'. \n"
                  f"Please check the following columns: {summary_check}")
            
        else:
            print(f"There are no differences between the summary of the gold standard and the input file for '{name}'.")

        detail_check = detail[detail != 0].dropna(how = 'all')

        if len(summary_check) > 0:
            print(f"There are differences between the detail of the gold standard and the input file for '{name}'. \n"
                  f"Please check the following columns: {detail_check}")
            
        else:
            print(f"There are no differences between the detail of the gold standard and the input file for '{name}'.")


if __name__ == '__main__':

    client_code = 7167
    fy = 2022

    if True:
        gold_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form2_formatted_{client_code}_{fy}.xlsx"
        input_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form2_formatted_{client_code}_{fy}.xlsx"

        form2 = MASForm2_Validator(gold_fp, input_fp)

    if True:
        gold_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form2_{client_code}_{fy}_awp.xlsx"
        input_fp = r"D:\workspace\luna\personal_workspace\gold_standard\mas_forms" + f"\mas_form2_{client_code}_{fy}_awp.xlsx"

        form2_awp = MASForm2_AWP_Validator(gold_fp, input_fp)