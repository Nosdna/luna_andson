import pandas as pd
import math
import logging
import os

# Initialise logger
logger = logging.getLogger()
if not(logger.hasHandlers()):
    logger.addHandler(logging.StreamHandler())

# Import luna package
import luna


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
            P = Form1Processor(df_to_process=data, luna_fp = self.luna_fp)

        df = P.process()
        return df

class FormProcessor:

    def __init__(self, df_to_process):
        self.df = df_to_process

    def process(self):
        
        self.remove_irrelevant_rows()
        self.extract_parenthesis_number()
        self.extract_amount()
        self.map_to_variable()

        return self.variables

    def remove_irrelevant_rows(self):

        if not hasattr(self, 'df_crop'):
            df = self.df.copy()

            df["strip_text"] = df["Alteryx OCR Output"].str.replace(" ", "")

            start_idx = df["strip_text"].str.contains(self.start_text, case=False, na=False).idxmax()

            end_idx = df["strip_text"].str.contains(self.end_text, case=False, na=False).idxmax()

            self.df_crop = df[start_idx:end_idx]

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
            
            amount_pattern = r'^(.*?)(-?\d[\-\?\d\,\_]*[\.\_\-\,\s][oOg\d]{2})[^0-9]*$'

            df[['Text', 'Amt2']] = df['Alteryx OCR Output'].str.extract(amount_pattern, expand=True)

            df[['Text', 'Amt1']] = df['Text'].str.extract(amount_pattern, expand=True)
            df = df.fillna('')

            # remove extracted amount value if the cell contains specific string
            df['Amt1'][df['strip_text'].str.contains('|'.join(self.rows_to_remove))] = ''
            df['Amt2'][df['strip_text'].str.contains('|'.join(self.rows_to_remove))] = ''

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
            
            df_cleaned['Amt1'] = df_cleaned['Amt1'].apply(convert_to_float)
            df_cleaned['Amt2'] = df_cleaned['Amt2'].apply(convert_to_float)

            self.output = df_cleaned.copy()

        return self.output
    
    def map_to_variable(self):
        
        df = self.extract_amount()
        no_of_amt = len(df[(df['Amt1'].notna())]) + len(df[(df['Amt2'].notna())])

        print(f"no_of_amt: {no_of_amt}, expected number of amt: {self.expected_number_of_vars}")

        if self.expected_number_of_vars != no_of_amt:
            raise Exception("Number of variables does not match expected number of variables")
        
        var_map = pd.read_excel(self.var_map_filepath, sheet_name="form1")
        
        list_of_amt = [val for val in df[['Amt1', 'Amt2']].values.flatten().tolist() if not math.isnan(val)]
        
        var_map["Amt"] = list_of_amt

        # Pivot the DataFrame
        var_map = var_map.reset_index()
        var_map["idx"] = var_map.index 
        var_map = var_map.pivot(index=['idx', 'var_name'], columns='amt_type', values='Amt')
        
        var_map = var_map.groupby(['var_name']).first().reset_index()

        # Remove the 'amt_type' label on top
        var_map.reset_index(inplace=True, drop=True)
        var_map = var_map.rename_axis(columns=None)

        self.variables = var_map


class Form1Processor(FormProcessor):
    
    
    def __init__(self, df_to_process, luna_fp):

        super().__init__(df_to_process=df_to_process)

        # self.var_map_filepath = "map_to_variable.xlsx"
        var_map_fn = "mas_f1_map_to_variable.xlsx"
        self.var_map_filepath = os.path.join(luna_fp, "parameters", var_map_fn)

        self.strings_to_replace = ["9.00", "9-00", "9-00", "9", "99", "9 99", "9 og"] # strings to replace with 0

        self.rows_to_remove = ["AnnualReturnv1.3",
                               "Datedthis(dd/mm/yy)",
                               "Statementofassetsandliabilitiesasat",
                               "ReportingCycle"]

        self.expected_number_of_vars = 130

        self.start_text = "SHAREHOLDERS’FUNDS/"

        self.end_text = "SupplementaryInformation"


if __name__ == "__main__":

    # # Get the luna folderpath 
    # luna_init_file = luna.__file__
    # luna_folderpath = os.path.dirname(luna_init_file)
    # logger.info(f"Your luna library is at {luna_folderpath}.")
    
    # # Get the template folderpath
    # template_folderpath = os.path.join(luna_folderpath, "templates")

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

    if True:

        # fp = r'data/Form1.xlsx'
        fp = r"D:\Documents\Project\Internal Projects\20231222 Code integration\OCR\fs-main\data\Form1.xlsx"

        # sheet_name = 'CrossInvest (Input)'
        # sheet_name = 'ICM Funds (Input)'
        sheet_name = 'Myer Gold (Input)'

        processor = OCROutputProcessor(filepath=fp, sheet_name=sheet_name, form="form1", luna_fp = settings.LUNA_FOLDERPATH)
        df = processor.execute()

        # output_fp = r'personal_workspace\output.xlsx'
        output_fp = r"D:\workspace\luna\personal_workspace\tmp\mas_f1_ocr_output.xlsx"
        df.to_excel(output_fp)