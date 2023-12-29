# Import standard libs
import os
import datetime
import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
import sys
sys.path.append("D:\workspace")


# Import luna package and fsvi package

import luna
import luna.common as common
import luna.fsvi as fsvi

from luna.common.gl import GLProcessor
import calendar

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

from luna.common.workbk import CreditDropdownList

class MASForm2_Generator_Part2:
    
    def __init__(self, tb_class, mapper_class,
                 gl_class, fy = 2022,
                 user_inputs = None
                 ):
        
        
        self.tb_class       = tb_class
        #self.aged_ar_class  = aged_ar_class
        self.mapper_class   = mapper_class
        self.fy             = fy
        self.gl_class       = gl_class
        self.user_inputs    = user_inputs

        self.main()
        
    def main(self):
        
        self.create_aa_tempdf()
        self.map_aa_df()        
        self.update_fees_receivables_tempdf(aged_receivables_fp)
        self.update_adjusted_asset_tempdf()
        self.map_adjusted_asset()
        self.map_avg_adjusted_asset()
        self.map_aa_threshold()
        self.column_mapper()
        self.output_excel()
        self.output_excel_trr()


#### (IV) Adjusted Assets 

    def create_aa_tempdf(self):
        # Adjusted Asset sections
        # Create 3 columns for last 3 month
        on_balance_assets_varname = "total_asset"
        on_balance_assets_row = self.mapper_class.varname_to_index.at[on_balance_assets_varname]

        aa_adjusted_assets_varname = "aa_adjusted_assets"
        aa_adjusted_assets_row = self.mapper_class.varname_to_index.at[aa_adjusted_assets_varname]
        
        aa_df = self.mapper_class.df_processed.copy()
        aa_df = aa_df.loc[on_balance_assets_row:aa_adjusted_assets_row]
        aa_df = aa_df.iloc[:, 0:5]


        aa_df.rename(columns={"Amount": self.col_name_3}, inplace=True) 
        col_names = aa_df.columns.tolist()
        col_names.insert(-1,self.col_name_2)
        col_names.insert(-1,self.col_name_1)
        aa_df = aa_df.reindex(columns = col_names)

        condition = aa_df.iloc[:,-1].notna() # change 999999 to Nan values in the last column 
        aa_df.iloc[condition, -1] = np.nan

        self.aa_df = aa_df

    def verify_credit_quality(self):
        
        tb_df = tb_class.get_data_by_fy(fy).copy()
        
        interval_list = [pd.Interval(5000,5000,closed = "both")]
        cash = tb_df[tb_df["L/S (interval)"].apply(lambda x:x in interval_list)]

        depo = cash[~cash["Name"].str.contains("(?i)cash")]
        
        new_cols = depo.columns.to_list() + ["Credit Quality Grade 1?"]
        
        depo = depo.reindex(columns=new_cols)

        
        print("Please tag if the following deposits are credit quality grade 1")
        
        depo.to_excel(r"D:\workspace\luna\personal_workspace\tmp\CreditQuality.xlsx")
        dropdownlist = CreditDropdownList(r"D:\workspace\luna\personal_workspace\tmp\CreditQuality.xlsx")
        depo = dropdownlist.create_dropdown_list()

        return depo
   
    def map_aa_df(self):
        month_end, first_month, second_month, third_month = self.gl_ageing()
        aa_df = self.aa_df.copy()
        print(aa_df)

        # On Balance sheet assets
        # from GL
        
        # Convert "L/S" column to numeric values
        # third_month["L/S"] = pd.to_numeric(third_month["L/S"], errors='coerce')  
        # second_month["L/S"] = pd.to_numeric(third_month["L/S"], errors='coerce')  
        # first_month["L/S"] = pd.to_numeric(third_month["L/S"], errors='coerce') 
         
        dec = third_month[third_month["L/S"]<6000]["Ending Balance"].sum()
        nov = second_month[second_month["L/S"]<6000]["Ending Balance"].sum()
        oct = first_month[first_month["L/S"]<6000]["Ending Balance"].sum()

        # pd.Interval(5000,6000, closed='left')

        on_balance_assets_varname = "total_asset"
        on_balance_assets_row = self.mapper_class.varname_to_index.at[on_balance_assets_varname]

        # Map to aa_df
        aa_df.at[on_balance_assets_row,self.col_name_1] = oct
        aa_df.at[on_balance_assets_row,self.col_name_2] = nov
        aa_df.at[on_balance_assets_row,self.col_name_3] = dec


        # Off balance sheet items 
        # Manual input
        answer  =  self.user_inputs.at["aa_off_bs_items", "Answer"]
        off_bal = answer.split(",")

        off_balance_assets_varname = "aa_off_bs_items"
        off_balance_assets_row = self.mapper_class.varname_to_index.at[off_balance_assets_varname]

        # Map to aa_df
        aa_df.loc[off_balance_assets_row, aa_df.columns[-3:]] = off_bal


        # Deductions from FR
        # from deductions_df
        fr_total_deductions_varname = "fr_total_deductions"
        fr_total_deductions_row = self.mapper_class.varname_to_index.at[fr_total_deductions_varname]

        total_deductions = self.deductions_df.loc[fr_total_deductions_row, self.deductions_df.columns[-3:]] 
        
        aa_fr_total_deductions_varname = "aa_fr_total_deductions"
        aa_fr_total_deductions_row = self.mapper_class.varname_to_index.at[aa_fr_total_deductions_varname]

        # Map to aa_df
        aa_df.loc[aa_fr_total_deductions_row, aa_df.columns[-3:]] = total_deductions


        # Cash and Deposit credit quality grade 1

        cash, depo = self._map_tempdf_by_ls([5000], "cash", type="Cash")
        print(cash)
        print(depo)

        aa_corp_own_cash_cashequiv_varname = "aa_corp_own_cash_cashequiv"
        aa_corp_own_cash_cashequiv_row = self.mapper_class.varname_to_index.at[ aa_corp_own_cash_cashequiv_varname]
        
        aa_df.loc[aa_corp_own_cash_cashequiv_row, aa_df.columns[-3:]] = cash

        aa_corp_own_deposit_bank_credit_quality_1_varname = "aa_corp_own_deposit_bank_credit_quality_1"
        aa_corp_own_deposit_bank_credit_quality_1_row = self.mapper_class.varname_to_index.at[aa_corp_own_deposit_bank_credit_quality_1_varname]

        aa_df.loc[aa_corp_own_deposit_bank_credit_quality_1_row, aa_df.columns[-3:]] = depo

        # Set values in aa_df to numeric 
        aa_df = aa_df.apply(pd.to_numeric, errors='ignore') 
        print(aa_df)
        print()
        print()
        print()
        print()
        print()
        

        self.aa_df = aa_df


#### AR processing for fees_receivables 
    def arprocessing(self, aged_receivables_fp, sheet_name):
        # Load the AR class
        aged_ar_class = common.AgedReceivablesReader_Format1(aged_receivables_fp, 
                                                        sheet_name = sheet_name, # Set the sheet name
                                                        variance_threshold = 0.1) # 1E-9) # To relax criteria if required.
        
        
        aged_group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                           ">90": ["91 - 120", "121 - 150", "150+"]}
        
        # Then we get the AR by company (index) and by new bins (columns)
        aged_df_by_company = aged_ar_class.get_AR_by_new_groups(aged_group_dict)
        self.aged_df_by_company = aged_df_by_company

        return self.aged_df_by_company
        
    def update_fees_receivables_tempdf(self, aged_receivables_fp):

        # To get the accounts receivables data for each month (already in SGD)
        thirdmonth = self.arprocessing(aged_receivables_fp,sheet_name = 0)
        secondmonth = self.arprocessing(aged_receivables_fp,sheet_name = 1)
        firstmonth = self.arprocessing(aged_receivables_fp, sheet_name = 2)

        # This company is non-CLMS, so need to remove from the computation
        answer = self.user_inputs.at["non_clms", "Answer"]
        non_clms = answer.split(",")

        # Filter away/remove non_clms for each month 
        thirdmonth_amt = thirdmonth[~thirdmonth.index.isin(non_clms)]["0-90"].sum()     #1660896.34
        secondmonth_amt = secondmonth[~secondmonth.index.isin(non_clms)]["0-90"].sum()  #65991.49
        firstmonth_amt = firstmonth[~firstmonth.index.isin(non_clms)]["0-90"].sum()     #1440337.46

        amt_list = [firstmonth_amt, secondmonth_amt, thirdmonth_amt]
        print(amt_list)
        print(type(amt_list[0]))

        # Less payment on behalf 
        answer = self.user_inputs.at["payment_on_behalf", "Answer"]
        paymt_behalf = answer.split(",") 

        print(paymt_behalf)
        print(type(paymt_behalf))
        paymt_behalf = [-float(x) for x in paymt_behalf] # convert to float and make it negative
        print(paymt_behalf)


        fees_receivables_varname="aa_fee_receivables_cis_cef_ca_within3mths"
        fees_receivables_row = self.mapper_class.varname_to_index.at[fees_receivables_varname]

        # Map payment on behalf & clms amt to aa_df
        self.aa_df.loc[fees_receivables_row, self.aa_df.columns[-3:]] = paymt_behalf

        for i, amt in enumerate(amt_list):
            self.aa_df.loc[fees_receivables_row, self.aa_df.columns[-3 + i]] += amt
    
        print(self.aa_df)
    
    def update_adjusted_asset_tempdf(self):
        
        # Get index rows for mapping later 
        aa_adjusted_assets_varname = "aa_adjusted_assets"
        aa_adjusted_assets_row = self.mapper_class.varname_to_index.at[aa_adjusted_assets_varname]
  
        off_balance_assets_varname = "aa_off_bs_items"
        off_balance_assets_row = self.mapper_class.varname_to_index.at[off_balance_assets_varname]

        aa_fr_total_deductions_varname = "aa_fr_total_deductions"
        aa_fr_total_deductions_row = self.mapper_class.varname_to_index.at[off_balance_assets_varname]

        fees_receivables_varname="aa_fee_receivables_cis_cef_ca_within3mths"
        fees_receivables_row = self.mapper_class.varname_to_index.at[fees_receivables_varname]

        print(aa_adjusted_assets_row)
        print(off_balance_assets_row)
        print(aa_fr_total_deductions_row)

      
        # Sum part 1(a) to (b)
        bal_items = self.aa_df.loc[:off_balance_assets_row, :].sum(axis = 0, numeric_only = True)  
        print(bal_items)

        # Sum part 1(c) to 1(f)
        remaining_items = self.aa_df.loc[aa_fr_total_deductions_row:fees_receivables_row, :].sum(axis = 0, numeric_only = True)
        print(remaining_items)

        # Adjusted assets = sum of 1(a,b) - sum of 1(c-f)
        total = bal_items - remaining_items
        print(total)
        
        # Map Adjusted asset amt to aa_df
        self.aa_df.loc[aa_adjusted_assets_row, self.aa_df.columns[-3:]] = total 

        # Compute Average adjusted asset 
        avg_adj_asset = np.mean(self.aa_df.loc[aa_adjusted_assets_row, self.aa_df.columns[-3:]])

        # Save as attr
        self.avg_adj_asset = avg_adj_asset

        print(avg_adj_asset)
        print(self.aa_df) 

    def map_adjusted_asset(self):

        aa_df = self.aa_df.copy()
                
        # Filter for the last month (Dec) amts 
        append_df = aa_df.iloc[:,-1]

        # Get index rows for mapping later 
        total_asset_varname = "total_asset"
        total_asset_row = self.mapper_class.varname_to_index.at[total_asset_varname]

        aa_adjusted_assets_varname = "aa_adjusted_assets"
        aa_adjusted_assets_row = self.mapper_class.varname_to_index.at[aa_adjusted_assets_varname]

        # Map to Template
        self.outputdf.loc[total_asset_row:aa_adjusted_assets_row, "Balance"] = append_df

    def map_avg_adjusted_asset(self):

        aa_avg_adjusted_assets_varname = "aa_avg_adjusted_assets"

        # Map amt to Template  
        self.add_bal_to_template_by_varname(aa_avg_adjusted_assets_varname, self.avg_adj_asset)

        # to delete
        print(f" Mapped {aa_avg_adjusted_assets_varname} : {self.avg_adj_asset}") 

    def map_aa_threshold(self):
        # Determine Adjusted Assets Threshold (5* TFR or $10M)
        # Get tfr from Template
        fr_varname = "fr_financial_resources_anhof"
        fr_row = self.mapper_class.varname_to_index.at[fr_varname]
        fr = self.outputdf.at[fr_row, "Balance"]

        print(fr)
        print(fr*5)
        threshold = min(5*(fr), 10000000)
        print(threshold)

        aa_adjusted_assets_threshold_varname = "aa_adjusted_assets_threshold"        
       
        # Map amt to Template  
        self.add_bal_to_template_by_varname(aa_adjusted_assets_threshold_varname, threshold)

        # to delete
        print(f" Mapped {aa_adjusted_assets_threshold_varname} : {threshold}") 

    def column_mapper(self):

        # Map the Balance amounts to the correct field in F1; whether in the Amount or Subtotal column
        for i in self.outputdf.index:
            if pd.notna(self.outputdf.at[i,"Amount"]) and pd.notna(self.outputdf.at[i+1,"Amount"]): #if the 2 rows in Amount column is not empty and row 1's Balance is not empty then compute subtotal
                if pd.isna(self.outputdf.at[i,"Balance"]):
                    subtotal = subtotal
                else:
                    subtotal += self.outputdf.at[i,"Balance"]
            elif pd.notna(self.outputdf.at[i,"Amount"]) and pd.isna(self.outputdf.at[i+1, #if the next row in Amount column is empty
                    "Amount"]):
                if pd.isna(self.outputdf.at[i,"Balance"]) and subtotal != 0: #if the Balance for a row is empty but the subtotal is not 0
                    self.outputdf.at[i,"Subtotal"] = subtotal
                    subtotal = 0
                elif pd.notna(self.outputdf.at[i,"Balance"]):      #if the Balance for a row is not empty 
                    subtotal += self.outputdf.at[i,"Balance"]
                    self.outputdf.at[i,"Subtotal"] = subtotal
                    subtotal = 0
            else: 
                subtotal = 0

        for i in self.outputdf.index:
            if pd.notna(self.outputdf.at[i, "Amount"]):                # if Amount column is not empty
                self.outputdf.at[i, "Amount"] = self.outputdf.at[i, "Balance"]
            elif pd.isna(self.outputdf.at[i, "Amount"]) and pd.notna(self.outputdf.at[i, # if Amount is empty and Subtotal not empty
                    "Subtotal"]):
                self.outputdf.at[i, "Subtotal"] = self.outputdf.at[i, "Balance"]
            elif pd.isna(self.outputdf.at[i, "Amount"]) and pd.isna(self.outputdf.at[i, # if Amount is empty and Subtotal not empty
                    "Subtotal"]):
                self.outputdf.at[i, "Subtotal"] = self.outputdf.at[i, "Balance"]  # if both Amount and Subtotal is empty 



#### Output for auditors (AAA)
    def output_excel(self):
        
        template_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\f2\Average Adjusted Assets Template.xlsx"
        
        aa_df = self.aa_df.copy()

        # To open the workbook 
        # workbook object is created
        wb = openpyxl.load_workbook(template_fp)
        
        # Get workbook active sheet object
        sheet = wb.active


        # first row (on-balance sheet assets)
        first_row = list(aa_df.iloc[0, -3:])

        # Inserting values into cells D9:F9
        columns = ['D', 'E', 'F']
        for col, value in zip(columns, first_row):
            sheet[col + '9'] = value
        
        # second row (off-balance sheet assets)
        second_row = list(aa_df.iloc[1, -3:])       
        # Inserting values into cells D11:F11
        for col, value in zip(columns, second_row):
            sheet[col + '11'] = value
        
        # To Less: 
        # third row (Cash & cash equivalents)
        third_row = list(aa_df.iloc[5, -3:])
        third_row = list(map(lambda x: -float(x), third_row)) # convert to float and make it negative
        #Inserting values into celss D13:f13
        for col, value in zip(columns, third_row):
            sheet[col + '13'] = value
        
        # fourth row (Deposits)
        fourth_row = list(aa_df.iloc[6, -3:])
        fourth_row = list(map(lambda x: -float(x), fourth_row)) # convert to float and make it negative
        #Inserting values into celss D14:f14
        for col, value in zip(columns, fourth_row):
            sheet[col + '14'] = value

        # fifth row (deductions from FR)
        fifth_row = list(aa_df.iloc[3, -3:])
        print(fifth_row)
        fifth_row = list(map(lambda x: -float(x), fifth_row)) # convert to float and make it negative
        #Inserting values into celss D15:f15
        for col, value in zip(columns, fifth_row):
            sheet[col + '15'] = value
        
        # sixth row (receivables owed by corporations)
        sixth_row = list(aa_df.iloc[7, -3:])
        sixth_row = list(map(lambda x: -float(x), sixth_row)) # convert to float and make it negative
        #Inserting values into celss D17:f17
        for col, value in zip(columns, sixth_row):
            sheet[col + '17'] = value

        # seventh row (receivables owed by corporations)
        seventh_row = list(aa_df.iloc[8, -3:])
        seventh_row = list(map(lambda x: -float(x), seventh_row)) # convert to float and make it negative
        #Inserting values into cells D18:f18
        for col, value in zip(columns, seventh_row):
            sheet[col + '18'] = value


        # Sum (a) to (iv) to get Adjusted asset 
        for col in range(4,7):
            formulas = "=SUM({}{}:{}{})".format(get_column_letter(col),
                                                9,
                                                get_column_letter(col),
                                                20
                                                )
            sheet.cell(row= 21 , column= col).value = formulas

        # Average adjusted asset 
        sheet['I21'] = "=AVERAGE(D21:F21)"

        # 5*FR
        fr_financial_resources_anhof_varname = "fr_financial_resources_anhof"
        fr_financial_resources_anhof_row = self.mapper_class.varname_to_index.at[fr_financial_resources_anhof_varname]
        print(fr_financial_resources_anhof_row)

        fr = self.outputdf.at[fr_financial_resources_anhof_row, "Balance"]

        print(fr)
        sheet['I26'] = float(fr)*5

        # Adjusted assets threshold 
        sheet['I29'] = "=MIN(I26,I27)"

        # Has AAA exceeded the above threshold?
        sheet['I31']= '=IF((I21<I26)*(I21<I27),"No","Yes")'
        
        wb.save("updated_workbook.xlsx")

#### Output for auditors (TRR)
    def output_excel_trr(self):
        template_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\f2\TRR.xlsx"

         # To open the workbook 
        # workbook object is created
        wb = openpyxl.load_workbook(template_fp)
        
        # Get workbook active sheet object
        sheet = wb.active

        # Retrieve from datahub (need current year)
        datahub = pd.read_excel("D:\gohjiawey\Desktop\Form 3\draftf2 - Copy.xlsx")
        datahub_currentfy = datahub[datahub['FY'] == self.fy]

    ###For current year

        # Base capital (row 8,9,10)        
        puc_list  = ["puc_ord_shares", "puc_pref_share_noncumulative", "puc_reserve_fund"]

        for i, row in enumerate(datahub_currentfy["var_name"]): 
            if row in (puc_list):
                value = datahub_currentfy.at[i, "Balance"]
                sheet["I"+ str(i+6)] = value #row 8,9,10

        # unappropriated profit or loss
        puc_unappr_profit_or_loss = datahub_currentfy.at[5, "Balance"]
        sheet["I12"] =  puc_unappr_profit_or_loss
        print(puc_unappr_profit_or_loss)

        # dividend declared, interim loss 
        puc_less_list = ["upl_div_declared", "upl_interim_loss"]
        for i, row in enumerate(datahub_currentfy["var_name"]): 
            if row in (puc_less_list):
                value = datahub_currentfy.at[i, "Balance"]
                sheet["I"+ str(i+8)] = value  #row 14,15

        # Net head office funds 
        puc_net_head_office_funds = datahub_currentfy.at[9, "Balance"]
        sheet["I17"] = puc_net_head_office_funds
        print(puc_net_head_office_funds)

        # Base Capital/Net Head Office Funds 
        sheet["I19"] =  "=SUM(I8:I17)"

        # Financial Resources 
        fr_add_list = ["puc_pref_share_cumulative", "current_liab_redeemable_pref_share, noncurrent_liab_redeemable_pref_share", 
         "bc_qual_subord_loans_temp", "puc_rev_reserve", "puc_other_reserves","bc_interim_unappr_profit",
         "bc_cltv_impairment_allowance"]
        
        for i, row in enumerate(datahub_currentfy["var_name"]):
            if row in (fr_add_list):
                value = datahub_currentfy.at[i, "Balance"]
                sheet["I"+ str(i+7)] = value #row 22 to 28

        
        fr_less_list = [ "noncurrent_asset_goodwill_ia", "dfr_future_incometax_benefits", "current_asset_other_prepayment", 
                "dfr_charged_assets", "dfr_unsecured_directors_cnntpersons", "dfr_unsecured_rlt_corporations",
                "dfr_other_unsecured_loans", "dfr_capinvst_subsidiary_associate", "noncurrent_asset_investment_in_subsi", 
                "dfr_other_assets_nonconvertible_cash"]

                     
        for i, row in enumerate(datahub_currentfy["var_name"]):
            if row in (fr_less_list):
                value = datahub_currentfy.at[i, "Balance"]
                sheet["I"+ str(i+5)] = -value #set to negative for less   #row 29 to 38 

        # Total Financial Resources ("FR")
        sheet["I40"] = "=SUM(I19:I38)"
        
        # ORR Highest of:
        sheet["I48"] = "=MAX(F50,F53)"

        # 5% of average annual gross income..
        sheet["F50"] = "=SUM(E92:E94)"

        # Total Risk Requirement ("TRR")
        sheet["I66"] = "=I48+I64"

        # Ratio: Financial Resources / Total Risk Requirement ("FR/TRR")
        sheet["I71"] = "=I40/I66"
                

    ###For last 3 years
        # Retrieve from datahub 
        datahub_form3 = pd.read_excel("D:\gohjiawey\Desktop\Form 3\draft_MG - Copy.xlsx")
        
        datahub_form3_current = datahub_form3[datahub_form3["FY"] == self.fy]

        required_varname_list = ["rev_total_revenue",
                                "exp_fee_expense", 
                                 "exp_comm_expense_otherbroker", 
                                 "exp_comm_expense_agents", 
                                 "exp_int_expense",
                                 "rev_int_others", 
                                 "rev_dividend", 
                                 "rev_other_revenue" 
                                 ]

        varname_dict = {}
        for i, row in enumerate(datahub_form3_current["var_name"]):
            if row in (required_varname_list):
                value = datahub_form3_current.at[i, "Previous Balance"]
                varname_dict[row] = value

        # Total Revenue        
        sheet["E78"] = varname_dict.get("rev_total_revenue", 0) # Default to 0 if key not present
        
        # Less Fees expense 
        sheet["E79"] = -varname_dict.get("exp_fee_expense", 0)

        # Less Commission expense 
        agents = -varname_dict.get('exp_comm_expense_agents', 0)  
        otherbroker = -varname_dict.get('exp_comm_expense_otherbroker', 0)  
        sheet["E80"] = agents + otherbroker
        
        # Less Interest expense 
        sheet["E81"] = -varname_dict.get("exp_int_expense", 0)

        # Less Income or expenses not derived from ordinary activities and not expected to recur frequently or regularly 
        answer = self.user_inputs.at["non_freq_income_exp", "Answer"]
        answer = answer.split(",")
        varname_list = []
        for i in answer:
            if re.search("(?i)div", i):
                varname_list.append("rev_dividend")
            elif re.search("(?i)other rev", i):
                varname_list.append("rev_other_revenue")
            elif re.search("(?i)interest"):
                varname_list.append("rev_int_others")

        income_exp_not_ord = 0
        for i in varname_list:
            income_exp_not_ord += varname_dict.get(i,0) 

        sheet["E83"] = -income_exp_not_ord  



        # Adjusted annual gross income 
        for col in range(3,6):
            formulas = "=SUM({}{}:{}{})".format(get_column_letter(col),
                                            78,
                                            get_column_letter(col),
                                            85
                                            )
            sheet.cell(row= 86 , column= col).value = formulas
            

        # Average annual gross income 
        sheet["E88"] = "=AVERAGE(C86,D86,E86)"

        # seperated into amts < S$10,000,000
        sheet["E89"] = "=MIN(E88,10000000)"

        # seperated into amts > S$10,000,000
        sheet["E90"] = "=MAX(0,E88-10000000)"

        # 5% average annual gross below S$10m
        sheet["E92"] = "=E89*0.05"

        # 2% of average annual gross income above S$10m
        sheet["E94"] = "=E90*0.02"

        # Save the workbook
        wb.save("updated_trr.xlsx")