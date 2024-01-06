import pandas as pd
import numpy as np
import math
import re
import os

class OCROutputProcessor:
    def __init__(self, filepath, sheet_name, form, luna_fp):
        self.filepath = filepath
        self.sheet_name = sheet_name
        self.form = form
        self.luna_fp = luna_fp
    
    def read_file(self):

        df_to_process = pd.read_excel(self.filepath, sheet_name=self.sheet_name)
        self.df_to_process = df_to_process

        return df_to_process

    def execute(self): # categorise, and process to generate the output
        
        data = self.read_file()

        if self.form == "form1":
            P = Form1Processor(df_to_process=data)
        elif self.form == "form2":
            P = Form2Processor(df_to_process=data, luna_fp = self.luna_fp)
        elif self.form == "form3":
            P = Form3Processor(df_to_process=data)
        else:
            raise Exception("Invalid form type")
        
        df = P.process()

        return df

class FormProcessor:

    def __init__(self, df_to_process, form, luna_fp):
        self.df = df_to_process
        self.form = form # to be used as sheet name
        var_map_fn = "map_to_variable.xlsx"
        no_of_vars_fn = "no_of_vars.xlsx"
        self.var_map_filepath = os.path.join(luna_fp, "parameters", var_map_fn)
        self.no_of_vars_filepath = os.path.join(luna_fp, "parameters", no_of_vars_fn)
        #self.var_map_filepath = "map_to_variable.xlsx"
        #self.no_of_vars_filepath = "no_of_vars.xlsx"

    def process(self):
        
        self.remove_irrelevant_rows()
        self.extract_parenthesis_number()
        self.extract_amount()
        self.validate_number_of_vars()
        self.map_to_variable()

        return self.variables

    def remove_irrelevant_rows(self):

        if not hasattr(self, 'df_crop'):
            df = self.df.copy()

            df["strip_text"] = df["Alteryx OCR Output"].str.replace(" ", "")

            def remove_rows(df, start_text, end_text, mode = "outside"):
                # mode = remove rows outside of start/end text
                # mode = remove rows between start/end text)

                start_idx = df["strip_text"].str.contains(start_text, case=False, na=False).idxmax()
                                
                end_idx = df["strip_text"].str.contains(end_text, case=False, na=False).idxmax()
                
                if mode == "outside":
                    
                    df_crop = df[start_idx:end_idx]

                elif mode == "between":
                    
                    df_crop = df.drop(index = range(start_idx+1,end_idx))

                return df_crop
            
            # remove rows before and after start text
            self.df_crop = remove_rows(df, 
                                       start_text= self.start_text, 
                                       end_text= self.end_text, 
                                       mode = "outside")

            # remove other revenue/expenses
            if self.form == "form3":
                    
                self.df_crop = remove_rows(self.df_crop, 
                                            start_text =self.other_rev_start_text, 
                                            end_text = self.other_rev_end_text, 
                                            mode = "between")
                
                self.df_crop = remove_rows(self.df_crop, 
                                            start_text=self.other_exp_start_text,
                                            end_text=self.other_exp_end_text, 
                                            mode = "between")

        return self.df_crop

    def extract_parenthesis_number(self):

        if not hasattr(self, 'df_processed'):

            # extract starting parenthesis from text col
            df = self.remove_irrelevant_rows()

            parenthesis_pattern = r'^(\([^)*]{1,2}\))'

            df['Number'] = df['Alteryx OCR Output'].replace(' ', '').str.extract(parenthesis_pattern, expand=True).fillna('')

            self.df_processed = df

        return self.df_processed
    
    def extract_amount(self):
        
        if not hasattr(self, 'output'):

            # extract numbers from text col
            df = self.extract_parenthesis_number()

            amount_pattern = r'^(.*?)(-?\d[\-\?\d\,\_]*[\,\.\_\-\s][oOg\d]{2})[^0-9]*$'
            amount_pattern2 = r'^(.*?)(-?\d[\-\?\d\,\_]*[\,\.\_\-\s][oOg\d]{2}).*$'

            df[['Text', 'Amt2']] = df['Alteryx OCR Output'].str.extract(amount_pattern, expand=True)
            df[['Text', 'Amt1']] = df['Text'].str.extract(amount_pattern, expand=True)

            # to handle 148.00
            df['Amt3'] = df['Alteryx OCR Output'].str.extract(amount_pattern2, expand=True)[1]
            mask = (df['Amt1'].isna()) & (df['Amt2'].isna()) & (~df['Amt3'].isna())
            df.loc[mask, 'Amt2'] = df.loc[mask, 'Amt3']
            df.drop(columns=['Amt3'], inplace=True)

            df = df.fillna('')

            for col in ['Amt1', 'Amt2']:
                df[col][df['strip_text'].str.match('|'.join(self.rows_to_remove))] = ''

            # format output
            df_cleaned = df[["Alteryx OCR Output", "Text", "Number", "Amt1", "Amt2"]].copy()
            
            # to check if Amt1/Amt2 in list of strings_to_replace
            df_cleaned[["Amt1", "Amt2"]] = df_cleaned[["Amt1", "Amt2"]].replace(self.strings_to_replace, "0.00")

            def convert_to_float(s):
                if s == '':
                    return None

                if s.endswith('.00'):
                    s = s[:-3]
                    s = s.replace(',', '')
                    return float(s)
                
                else:
                    s = s.replace(',', '')
                    return float(s)
            
            for col in ['Amt1', 'Amt2']:
                df_cleaned[col] = df_cleaned[col].apply(convert_to_float)

            self.output = df_cleaned.copy()
        
        return self.output

    def validate_number_of_vars(self):

        if self.get_no_of_missing_vars() == 0:
            return

        if not hasattr(self, 'df_with_missing_tag'):
            
            df = self.extract_amount()

            # get parenthesis numbers present in OCR result, and update no_of_var
            no_of_var = self.match_parenthesis_number(df)

            # for present parenthesis numbers, create comparison df
            comp_df = self.create_comparison_df(target_df=no_of_var, ocr_df=df)

            # identify the section(s) with missing values
            self.df_with_missing_tag = self.identify_missing_values_by_level(comparison_df=comp_df)
            
        return self.df_with_missing_tag
        
    def get_no_of_missing_vars(self):
        
        if not hasattr(self, 'no_of_missing_vars'):

            df = self.extract_amount()

            no_of_amt = len(df[(df['Amt1'].notna())]) + len(df[(df['Amt2'].notna())])

            print(f"no_of_amt: {no_of_amt}, expected number of amt: {self.expected_number_of_vars}")

            self.no_of_missing_vars = self.expected_number_of_vars - no_of_amt

        return self.no_of_missing_vars
    
    def match_parenthesis_number(self, df):

        # get excel row idx
        df['ridx'] = df.index + 2

        # exclude rows with no parenthesis number
        df_numbers = df[df['Number'] != '']['Number']

        # get expected parenthesis numbers
        no_of_var = pd.read_excel(self.no_of_vars_filepath, sheet_name=self.form)

        expected_numbers = no_of_var['number'].unique().tolist()

        # match numbers present in OCR result with expected parenthesis numbers
        adjustment = 0

        for i in range(len(no_of_var)):
            
            value_to_match = df_numbers.iloc[i+adjustment].replace(' ', '')
            # print(value_to_match)
            # print(f"i:{i}, adjustment:{adjustment}, i+adjustment:{i+adjustment}")

            if re.search(r'\([^0-9a-zA-Z]+\)', value_to_match):
                # to handle incorrectly read numbers like (@)
                df.loc[df['Number'] == value_to_match, 'Number'] = ''
                # print("> skipped (special char)")
                continue

            elif value_to_match not in expected_numbers:
                # if the value is read as wrong alphanumeric that is present in no_of_vars, not working 
                #  e.g. (II) > (ID): OK
                #       (a) > (b): X
                df.loc[df['Number'] == value_to_match, 'Number'] = ''
                # print("> skipped (not in expected numbers)")
                continue

            elif no_of_var.iloc[i]["number"] == value_to_match:
                no_of_var.loc[i, "numbers_present"] = no_of_var.loc[i,"number"]
                # print("> matched")

            else:
                adjustment -= 1
                # print("> searching")

        no_of_var = no_of_var.replace('nan', '')

        return no_of_var
    
    def create_comparison_df(self, target_df, ocr_df):

        def group_by_num_present(df, number_present_col: str, agg_dict: dict):
            # forward-fill number col
            df['Number_filled'] = df[number_present_col].replace('',np.nan).ffill()

            # generate running number
            df['group'] = (df['Number_filled'] != df['Number_filled'].shift()).cumsum()

            #set agg function for number col and group df
            agg_dict["Number_filled"] = "first"

            return df.groupby('group').agg(agg_dict)
        
        # get expected number of vars for each parenthesis number present in OCR result
        target_no_of_var = group_by_num_present(df=target_df,
                                                number_present_col='numbers_present',
                                                agg_dict={'no_of_vars':'sum',
                                                          'number_lvl':'first'})

        # get count of amt present per num in OCR result
        df_no_of_var = group_by_num_present(df=ocr_df,
                                            number_present_col='Number',
                                            agg_dict= {'Amt1':'count',
                                                       'Amt2':'count'})

        df_no_of_var['no_of_var_present'] = df_no_of_var['Amt1'] + df_no_of_var['Amt2']

        # get df comparing amt present/expected
        comparison_df = target_no_of_var[['Number_filled', 'number_lvl', 'no_of_vars']].copy()
        comparison_df['no_of_var_present'] = df_no_of_var['no_of_var_present']

        return comparison_df

    def identify_missing_values_by_level(self, 
                                         comparison_df, 
                                         start_level=1, 
                                         end_level=4):
        result_dict = {}

        new_col_name = f"level{start_level}"

        for i in range(start_level, end_level+1): # (inclusive)

            if not i == start_level:
                new_col_name += f'+{i}'

            # extract parenthesis numbers for each level as new col
            comparison_df[new_col_name] = np.where(comparison_df['number_lvl'] <= i,
                                                   comparison_df['Number_filled'],
                                                   '')
            comparison_df[new_col_name] = (comparison_df[new_col_name]
                                           .replace('', None)
                                           .ffill())

            # group col will be used as groupby key col
            group_col_name = f"group{i}"
            comparison_df[group_col_name] = (comparison_df[new_col_name]
                                             != comparison_df[new_col_name].shift(1)
                                             ).cumsum()
            
            # for each level, get sum of expected and present no_of_vars
            comparison_df[f'target_lv{i}_sum'] = (
                comparison_df
                .groupby(group_col_name)['no_of_vars']
                .transform('sum')
                )
            
            comparison_df[f'present_lv{i}_sum'] = (
                comparison_df
                .groupby(group_col_name)['no_of_var_present']
                .transform('sum')
                )
            
            comparison_df.reset_index(inplace=True, drop=True)
            comparison_df['idx'] = comparison_df.index
            comparison_df['missing_section'] = False

            # compare calculated no_of_vars, calculate diff for each level
            target_dict = (comparison_df
                           .groupby(group_col_name)[f'target_lv{i}_sum']
                           .first().to_dict())
            present_dict = (comparison_df
                           .groupby(group_col_name)[f'present_lv{i}_sum']
                           .first().to_dict())
            
            result_dict[i] = {g:target_dict[g]-present_dict[g] 
                              for g in comparison_df[group_col_name].to_list()}

        grouped1_present = comparison_df.groupby('group1')
        
        for group1_val, diff in result_dict[1].items():
            if diff != 0:
                df = grouped1_present.get_group(group1_val).copy()
                grouped2_present = df.groupby('group2')

                group2_values = df['group2'].to_list()

                for group2_val in group2_values:
                    if result_dict[2][group2_val] != 0: # diff
                        df = grouped2_present.get_group(group2_val).copy()
                        grouped3_present = df.groupby('group3')

                        group3_values = df['group3'].to_list()

                        for group3_val in group3_values:
                            if result_dict[3][group3_val] != 0: # diff
                                df = grouped3_present.get_group(group3_val).copy()
                                
                                comparison_df.loc[df['idx'], 'missing_section'] = True

        comparison_df['missing_section_diff'] = comparison_df.apply(
            lambda row: result_dict[4].get(row['group4']) if row['missing_section'] else None, axis=1)

        comparison_df = comparison_df[['Number_filled', 'group4', 
                                       'no_of_vars', 'no_of_var_present', 
                                       'level1', 'target_lv1_sum', 'present_lv1_sum',
                                       'level1+2', 'target_lv2_sum', 'present_lv2_sum',
                                       'level1+2+3', 'target_lv3_sum', 'present_lv3_sum',
                                       'level1+2+3+4', 'target_lv4_sum', 'present_lv4_sum',
                                       'missing_section', 'missing_section_diff']].copy()
        

        self.comparison_df = comparison_df
        self.result_dict = result_dict
        
        return self.comparison_df
    
    def identify_missing_values(self, df): # abi 
        # calculate cumulative sum for expected and present amts
        df["cum_no_of_var"] = df["no_of_vars"].cumsum()
        df["cum_no_of_var_present"] = df["no_of_var_present"].cumsum()

        df["value_matches"] = df["no_of_vars"] == df["no_of_var_present"]

        missing_section = False
        for index in df.index:
            row = df.loc[index]

            if not row['value_matches']:

                # flag as error only if missing cumulatively
                if row['cum_no_of_var_present'] < row['cum_no_of_var']:
                    missing_section = True
                    df.loc[index, 'missing_section'] = True

                # flag as error UNTIL cumulative value matches
                if missing_section:
                    if not row['value_matches']:
                        print("MISS")
                        df.loc[index, 'missing_section'] = True
            
            # restart calculation of cumulative col once value matches
            elif missing_section and row['value_matches']: 
                df.loc[index:, 'cum_no_of_var'] = df.loc[index:, 'no_of_vars'].cumsum()
                df.loc[index:, 'cum_no_of_var_present'] = df.loc[index:, 'no_of_var_present'].cumsum()
                missing_section = False

        return df
    
    def interpolate_missing_values(self, df):
        
        if not hasattr(self, 'comparison_df'):
            return None, None
        
        comparison_df = self.comparison_df

        merged_df = df.merge(comparison_df[['group4', 'missing_section', 'missing_section_diff']], 
                             left_on='group', right_on='group4', how='left'
                             ).drop(columns=['group'])
        merged_df['missing_group'] = (merged_df['missing_section']
                                      .ne(merged_df['missing_section'].shift())
                                      .cumsum())
        
        new_list_of_amt = []
        new_missing_identifier = []
        for chunk in merged_df['missing_group'].unique():
            chunk_df = merged_df[merged_df['missing_group']==chunk]
            is_missing_section = chunk_df['missing_section'].all()

            if not is_missing_section: # if not missing section, just append
                present_var = [val 
                               for val 
                               in chunk_df[['Amt1', 'Amt2']].values.flatten().tolist()
                               if not math.isnan(val)]
                new_list_of_amt += present_var
                new_missing_identifier += [False] * len(present_var)
            
            else: # if missing section, all values will be converted to -9999.99
                to_add = int(chunk_df.groupby('group4')['missing_section_diff'].first().sum())
                present_var = [val 
                               for val 
                               in chunk_df[['Amt1', 'Amt2']].values.flatten().tolist()
                               if not math.isnan(val)]
                
                list_to_add = [-9999.99] * (len(present_var))
                identifier_to_add = [True] * (len(present_var))

                if to_add < 0: # provided more than required
                    new_list_of_amt += list_to_add[:to_add]
                    new_missing_identifier += identifier_to_add[:to_add]

                else: # missing values
                    new_list_of_amt += (list_to_add + [-9999.99]*to_add)
                    new_missing_identifier += (identifier_to_add + [True]*to_add)

        return new_list_of_amt, new_missing_identifier

        # return merged_df

    def map_to_variable(self):

        df = self.extract_amount()
        
        list_of_amt = [val 
                       for val 
                       in df[['Amt1', 'Amt2']].values.flatten().tolist() 
                       if not math.isnan(val)]
        
        print(list_of_amt)

        missing_identifier = False

        new_list_of_amt, new_missing_identifier = self.interpolate_missing_values(df=df)

        if new_list_of_amt is not None: # if there is missing section, use the interpolated list
            list_of_amt = new_list_of_amt
            missing_identifier = new_missing_identifier

        no_of_amt = len(list_of_amt)

        if self.expected_number_of_vars != no_of_amt:
            print("LIST OF AMT HERE")
            print(list_of_amt)
            print(no_of_amt)
            print(f"no_of_amt: {no_of_amt}, expected number of amt: {self.expected_number_of_vars}")
            raise Exception("Number of variables does not match expected number of variables")
        
        var_map = pd.read_excel(self.var_map_filepath, sheet_name=self.form)
        var_map["Amt"] = list_of_amt
        var_map["missing_identifier"] = missing_identifier

        # Pivot the DataFrame
        var_map = var_map.reset_index()
        var_map["idx"] = var_map.index 
        var_map = var_map.pivot(index=['idx', 'var_name', 'missing_identifier'], 
                                columns='amt_type', 
                                values='Amt').reset_index()

        var_map = (var_map.groupby(['var_name'])
                   .first()
                   .reset_index()
                   .sort_values(by='idx')
                   .reset_index())
        
        self.variables = var_map[['var_name', 'amount', 
                                  'subtotal', 'missing_identifier']].copy()
        self.variables.to_excel("var_map.xlsx")


class Form1Processor(FormProcessor):
    
    def __init__(self, df_to_process, form="form1"):

        super().__init__(df_to_process=df_to_process, form=form)

        self.strings_to_replace = ["9.00", "9-00", "9-00", "9", "99", "9 99", "9 og"] # strings to replace with 0

        self.rows_to_remove = ["AnnualReturnv1.3",
                               "Datedthis(dd/mm/yy)",
                               "Statementofassetsandliabilitiesasat",
                               "ReportingCycle"]

        self.expected_number_of_vars = 130

        self.start_text = "NETHEADOFFICEFUNDS"

        self.end_text = "SupplementaryInformation"

class Form2Processor(FormProcessor):
    
    def __init__(self, df_to_process, luna_fp, form="form2"):

        super().__init__(df_to_process=df_to_process, luna_fp = luna_fp, form=form)

        self.strings_to_replace = ["9.00", "9-00", "9-00", "9", "99", "9 99", "9 og"] # strings to replace with 0

        self.rows_to_remove = ["AnnualReturnv1.3",
                               "Datedthis(dd/mm/yy)",
                               "Statementofassetsandliabilitiesasat",
                               "ReportingCycle"]

        self.expected_number_of_vars = 94 # number of "amounts" not vars (some amts no var name..)

        self.start_text = "NETHEADOFFICEFUNDS"

        self.end_text = "Statementbyholderofcapital"

class Form3Processor(FormProcessor):
    
    def __init__(self, df_to_process, form="form3"):

        super().__init__(df_to_process=df_to_process, form=form)

        self.strings_to_replace = ["9.00", "9-00", "9-00", "9", "99", "9 99", "9 og"] # strings to replace with 0

        self.rows_to_remove = ["AnnualReturnv1.3",
                               "Datedthis(dd/mm/yy)",
                               "Statementofassetsandliabilitiesasat",
                               "ReportingCycle",
                               "20__",
                               "Significant if amount is greater or equal to 5% of Total Revenue.",
                               "Significant if amount is greater or equal to 5% of Total Expenses"]


        self.expected_number_of_vars = 88

        self.start_text = r"^20__"

        self.end_text = "Statementbyholderofcapital"

        self.other_rev_start_text = "Otherrevenue"
        self.other_rev_end_text = r"^TotalRevenue"

        self.other_exp_start_text = "Otherexpenses"
        self.other_exp_end_text = r"^TotalExpenses"

if __name__ == "__main__":

    if True:
        import importlib.util
        loginid = os.getlogin().lower()
        if loginid == "owghimsiong":
            settings_py_path = r'D:\Desktop\owgs\CODES\luna\settings.py'
        elif loginid == "phuasijia":
            settings_py_path = r'D:\workspace\luna\settings.py'
        else:
            raise Exception (f"Invalid user={loginid}. Please specify the path of settings.py.")
        
        # Import the luna environment through settings.
        # DO NOT TOUCH
        spec = importlib.util.spec_from_file_location("settings", settings_py_path)
        settings = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(settings)

    # from reader import OCROutputProcessor
    # from reader import Form3Processor
    # fp1 = r'abi/data/Form1.xlsx'
    # fp2 = r'abi/data/Form2.xlsx'
    # fp3 = r'abi/data/Form3.xlsx'

    fp2 = r"D:\Documents\Project\Internal Projects\20231222 Code integration\fs for sijia\data\Form2.xlsx"

    # sheet_name = 'CrossInvest (Input)'
    sheet_name = 'ICM Funds (Input)' 
    # sheet_name = 'Myer Gold (Input)'

    processor = OCROutputProcessor(filepath=fp2, sheet_name=sheet_name, form="form2", luna_fp = settings.LUNA_FOLDERPATH)
    df = processor.execute()


    output_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_f2_ocr_output.xlsx"
    df.to_excel(output_fp)

