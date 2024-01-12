'''
Sample script for Dacia and Jia Wey
'''

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

class MASForm1_Generator:

    def __init__(self, 
                 tb_class, aged_ar_class, mapper_class,
                 client_class, ocr_class, fy,
                 fuzzy_match_threshold = 80, 
                 user_inputs = None):
        
        self.tb_class               = tb_class
        self.aged_ar_class          = aged_ar_class
        self.mapper_class           = mapper_class
        self.client_class           = client_class
        self.ocr_class              = ocr_class
        self.fy                     = fy
        self.fuzzy_match_threshold  = fuzzy_match_threshold
        self.user_inputs            = user_inputs
        
        self.main()
        
    def main(self):
        
        # Will extract the tb and mapper of varname to lscodes
        self._map_varname_to_lscodes()

        # Prepare output container -> copy from mapper_class
        self._prepare_output_container()

        # If user_inputs not specified, collect manual inputs interactively
        if self.user_inputs is None:
                            
            self.collect_manual_inputs()
                                
        else:
            pass

        # Map TB numbers to outputdf
        self.map_tb_to_output()
        
        # Ordinary Shares
        self._calculate_field_ord_share()

        # Preference Share - Irredeemable Cumulative
        self._calculate_field_pref_share_irredeemable_cumulative()

        # Preference Share - Irredeemable Non-Cumulative
        self._calculate_field_pref_share_irredeemable_noncumulative()

        # Preference Share - Redeemable Current
        self._calculate_field_pref_share_redeemable_current()

        # Preference Share - Redeemable Non-current
        self._calculate_field_pref_share_redeemable_noncurrent()

        # Trade creditors
        self._calculate_field_trade_cred()

        # Amount due to director
        self._calculate_field_amt_due_to_dir()

        # Loan from related corporations
        self._calculate_field_loans_from_relatedco()

        # Other current liability
        self._calculate_field_other_current_liab()

        # Trade debtors - fund management
        self._calculate_field_trade_debtors_fundmgmt(self.fuzzy_match_threshold)

        # Trade debtors - others
        self._calculate_field_trade_debt_other()

        # Amount due from director - Secured
        self._calculate_field_amt_due_from_dir_secured()

        # Amount due from director - Unsecured
        self._calculate_field_amt_due_from_dir_unsecured()

        # Loan to related corporations
        self._calculate_field_loan_to_relatedco()

        # Other current assets - Deposit
        self._calculate_field_deposit()
        
        # Other current assets - Prepayment
        self._calculate_field_prepayment()

        # Other current assets - Others
        self._calculate_field_other_current_asset()

        # Absolute of Balance except for unappropriate profit or loss
        self.abs_of_balance_column()

        # Map balances to the correct column (Amount or Subtotal column) & calculate amount totals to append in subtotal column
        self.column_mapper()

        # Calculate row totals
        self.get_row_totals()
        
        # Load output to lunahub
        self.load_output_to_lunahub()

        # Process OCR output
        self.process_ocr_output()


    def _prepare_output_container(self):

        # make a copy of the input template
        outputdf = self.mapper_class.df_processed.copy()

        # Add an output column
        outputdf["Balance"] = None

        # Save as attr
        self.outputdf = outputdf

    def _map_varname_to_lscodes(self):
        
        mapper_class = self.mapper_class
        tb_class     = self.tb_class
        
        # get varname to ls code from mapper
        varname_to_lscodes = mapper_class.varname_to_lscodes

        # NOTE: edited by SJ to accommodate formulas on mapping template 20240111
        varname_to_lscodes_temp = varname_to_lscodes.copy()
        varname_to_lscodes_temp = varname_to_lscodes_temp.to_frame()
        pattern = ".*Interval.*"
        varname_to_lscodes_temp["filter"] = np.where(varname_to_lscodes_temp["L/S (intervals)"].astype(str).str.match(pattern), "yes", "no")
        varname_to_lscodes_ls = varname_to_lscodes_temp[varname_to_lscodes_temp["filter"] == "yes"]["L/S (intervals)"].squeeze()
        varname_to_lscodes_formula = varname_to_lscodes_temp[varname_to_lscodes_temp["filter"] == "no"]["L/S (intervals)"].squeeze()
        
        varname_to_lscodes = varname_to_lscodes_ls
        
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
        
        varname_to_lscodes = pd.concat([varname_to_lscodes, varname_to_lscodes_formula], axis = 0)

        #
        self.tb_columns_main = tb_columns
        self.tb_with_varname = tb_df.copy()
        self.tb_main = tb_df[tb_columns]

    def filter_tb_by_varname(self, varname):
        
        tb = self.tb_with_varname.copy()
        
        filtered_tb = tb[tb[varname]][self.tb_columns_main]
        
        return filtered_tb

    def _get_total_by_varname(self, varname):
        
        # Get interval list from varname (template) 
        interval_list = self.mapper_class.get_ls_codes_by_varname(varname)

        try: # NOTE: Added by SJ on 20240111

            # Get the TB accounts that overlap in the interval list
            boolean, true_match, false_match = \
                self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, interval_list)
            
            # Return the total
            total = true_match["Value"].sum()
        
        except:

            print(f"Could not convert to interval for {varname}.")
            total = 0
        
        return total

    def map_tb_to_output(self):
        # Loop through the varnames in the template and append the balances from the tb
        varname_df = self.outputdf["var_name"].dropna()

        

        self.outputdf["Balance"] = varname_df.apply(self._get_total_by_varname)


    def add_bal_to_template_by_varname(self, varname, value, is_total_row=None):
        """
        total -> only need to use if it is a total row, the amount will be added to the subtotal column

        if not a total, amount added to the balance column
        """
        if is_total_row is None:
            row = self.mapper_class.varname_to_index.at[varname]
            self.outputdf.loc[row, "Balance"] = value

        else:
            row = self.mapper_class.varname_to_index.at[varname]
            self.outputdf.loc[row, "Subtotal"] = value
    

    def _auto_tagging_for_shares(self, filtered_tb):
        """
        Classify shares into 5 categories by regular matching of patterns:

        Irredeemable preference shares will be split into cumulative and non-cumulative, if no pattern matched for cumulative or not, the default is set as cumulative
        (a) Irredeemable Preference Share (Cumulative)
        (b) Irredeemable Preference Share (Non-Cumulative)

        Redeemable preference shares will be split into current and non-current portions, if no pattern matched for current or not, the default is set as current
        (c) Redeemable Preference Share (Current)
        (d) Redeemable Preference Share (Non-Current)

        All remaining untagged accounts are classified as ordinary shares
        (e) Ordinary Shares

        """
        for i in filtered_tb.index:
            # Search for Irredeemable Preference Shares
            if re.search("(?i)irrede.*pref.*", filtered_tb.at[i, "Name"]):
                # Search for non-cumulative in irredeemable preference shares
                if re.search("(?i)non.*cumu.*", filtered_tb.at[i, "Name"]):
                    filtered_tb.at[i, "Tag"] = "Irredeemable Preference Share (Non-Cumulative)"
                # If account name has no tagging of cumulative or non-cumulative, default set as cumulative irredeemable preference share
                else:
                    filtered_tb.at[i, "Tag"] = "Irredeemable Preference Share (Cumulative)"

            # Search for Redeemable Preference Shares
            elif re.search("(?i)(?<!ir)redeem.*pref.*", filtered_tb.at[i, "Name"]): 
                # Search for non-current in redeemable preference shares
                if re.search("(?i)non.*current", filtered_tb.at[i, "Name"]):
                    filtered_tb.at[i, "Tag"] = "Redeemable Preference Share (Non-Current)"
                # If account name has no tagging of current or non-current, default set as current redeemable preference share
                else: 
                    filtered_tb.at[i, "Tag"] = "Redeemable Preference Share (Current)"
            # All remaining untagged accounts to tag as ordinary shares
            else:
                filtered_tb.at[i, "Tag"] = "Ordinary Shares"

        return filtered_tb
    

    def _calculate_field_ord_share(self):
        """
        Filter tb by varname for ord_shares, return value of ordinary shares
        Append value to self.outputdf
        """
        varname = "puc_ord_shares"
        filtered_tb = self.filter_tb_by_varname(varname)
        filtered_tb["Value"].sum()

        if not hasattr(self, 'tagged_shares'):
            tagged_shares = self._auto_tagging_for_shares(filtered_tb)
        
            # save as attr
            self.tagged_shares = tagged_shares

        print(self.tagged_shares)

        # Query for Ordinary Shares
        ord_share = self.tagged_shares.query("Tag=='Ordinary Shares'")
        ord_share = ord_share["Value"].sum()
        self.add_bal_to_template_by_varname(varname, ord_share)


    def _calculate_field_pref_share_irredeemable_cumulative(self):
        varname = "puc_pref_share_cumulative"

        cumulative_share = self.tagged_shares.query("Tag=='Irredeemable Preference Share (Cumulative)'")
        cumulative_share = cumulative_share["Value"].sum()
        self.add_bal_to_template_by_varname(varname, cumulative_share)
    
    
    def _calculate_field_pref_share_irredeemable_noncumulative(self):
        varname = "puc_pref_share_noncumulative"

        noncumulative_share = self.tagged_shares.query("Tag=='Irredeemable Preference Share (Non-Cumulative)'")
        noncumulative_share = noncumulative_share["Value"].sum()
        self.add_bal_to_template_by_varname(varname, noncumulative_share)


    def _calculate_field_pref_share_redeemable_current(self):
        varname = "current_liab_redeemable_pref_share"

        redeemable_current_share = self.tagged_shares.query("Tag=='Redeemable Preference Share (Current)'")
        redeemable_current_share = redeemable_current_share["Value"].sum()
        self.add_bal_to_template_by_varname(varname, redeemable_current_share)


    def _calculate_field_pref_share_redeemable_noncurrent(self):
        varname = "noncurrent_liab_redeemable_pref_share"

        redeemable_noncurrent_share = self.tagged_shares.query("Tag=='Redeemable Preference Share (Non-Current)'")
        redeemable_noncurrent_share = redeemable_noncurrent_share["Value"].sum()
        self.add_bal_to_template_by_varname(varname, redeemable_noncurrent_share)


    def _calculate_field_deposit(self):
        """
        Classify shares into 3 categories by regular matching of patterns:

        (a) Deposit
        (b) Prepayment
        (c) Others

        If no pattern matched, classified as others
        """
        varname = "current_asset_other_deposit"
        filtered_tb = self.filter_tb_by_varname(varname)

        if not hasattr(self, 'depo_prepaid'):

            depo_prepaid = filtered_tb.copy()

            for i in depo_prepaid.index:
                if re.search("(?i)pre", depo_prepaid.at[i, "Name"]):
                    depo_prepaid.at[i, "Indicator"] = "Prepayment"
                elif re.search("(?i)deposit", depo_prepaid.at[i, "Name"]):
                    depo_prepaid.at[i, "Indicator"] = "Deposit"
                else:
                    depo_prepaid.at[i, "Indicator"] = "Others"
        
            # save as attr
            self.depo_prepaid = depo_prepaid

        deposit = self.depo_prepaid.query("Indicator=='Deposit'")
        deposit = deposit["Value"].sum()

        self.add_bal_to_template_by_varname(varname, deposit)


    def _calculate_field_prepayment(self):
        varname = "current_asset_other_prepayment"

        prepayment = self.depo_prepaid.query("Indicator=='Prepayment'")
        prepayment = prepayment["Value"].sum()

        self.add_bal_to_template_by_varname(varname, prepayment)


    def _calculate_field_other_current_asset(self):
        varname = "current_asset_other_other"

        others = self.depo_prepaid.query("Indicator=='Others'")
        others = others["Value"].sum()

        self.add_bal_to_template_by_varname(varname, others)

        # Need to minus loans and amt due from dir
        varname_list = ["current_asset_amount_due_from_director_secured", 
                        "current_asset_amount_due_from_director_unsecured", 
                        "current_asset_loans_to_related_co"]
        rows = self.mapper_class.varname_to_index.loc[varname_list]
        rpt_asset = self.outputdf.loc[rows, "Balance"].sum()

        others_row = self.mapper_class.varname_to_index.at[varname]
        self.outputdf.loc[others_row, "Balance"] -= rpt_asset


    def _calculate_field_trade_debtors_fundmgmt(self, fuzzy_match_threshold):
        """
        fuzzy_match_threshold -> int to set threshold for fuzzy matching of client/supplier names to the list provided by auditor
        """
        varname = "current_asset_trade_debt_fund_mgmt"
        
        if self.aged_ar_class is None:
            
            # No AR available
            self.add_bal_to_template_by_varname(varname, 0)
            
        else:

            # AR is available for the client
            self.fy_end_month = self.client_class.retrieve_fy_end_date(self.fy).month
            ar = self.aged_ar_class.df_processed_long_lcy
            ar = ar[ar["Month"] == self.fy_end_month]
            ar = ar[["Name", "Value (LCY)"]]
            
            # Convert input string to list and strip the leading/trailing spaces
            answer = self.user_inputs.at["trade_debt_fund_mgmt", "Answer"]
    
            fundmgmt_list = [i.strip() for i in answer.split(",")]
    
            ar.loc[:,"Match_score"] = ar["Name"].apply(lambda x: process.extractOne(x, fundmgmt_list, scorer=fuzz.token_sort_ratio))
    
            # Indicate if there is a match, match if 'Match_score' score is above fuzzy_match_threshold provided
            ar.loc[:,"Matched?"] = ar["Match_score"].apply(lambda x: x[1]) >= fuzzy_match_threshold
    
            # Filter for matches and obtain the sum of Total Due
            fundmgmt_df = ar.query("`Matched?`==True")
    
            fundmgmt_debtors = fundmgmt_df["Value (LCY)"].sum()
    
            self.add_bal_to_template_by_varname(varname, fundmgmt_debtors)


    def _calculate_field_trade_debt_other(self):
        # Minus amount of trade debt for fund management from total trade debt to get others
        fundmgmt_row = self.mapper_class.varname_to_index.at["current_asset_trade_debt_fund_mgmt"]
        fund_mgmt_debtors = self.outputdf.loc[fundmgmt_row, "Balance"]

        others_row = self.mapper_class.varname_to_index.at["current_asset_trade_debt_other"]
        self.outputdf.loc[others_row, "Balance"] -= fund_mgmt_debtors


    def _calculate_field_trade_cred(self):
        """
        For now, collecting manual inputs for $ amount of total trade creditors and fund management

        Agreed treatment: Fuzzy matching of client account names related to fund management 
        (Can use the fuzzy matching from calculate_field_trade_debtors_fundmgmt above)
        """ 

        # Trade creditor for fund management
        varname = "current_liab_trade_cred_fund_mgmt"
        
        answer = self.user_inputs.at["trade_cred_fund_mgmt", "Answer"]
        fundmgmt_cred = -int(float(answer))
        self.add_bal_to_template_by_varname(varname, fundmgmt_cred)


        # Other trade creditor
        answer = self.user_inputs.at["total_trade_cred", "Answer"]
        total_trade_cred = -int(float(answer))
        other_trade_cred = total_trade_cred-fundmgmt_cred
        varname = "current_liab_trade_cred_other_other"
        self.add_bal_to_template_by_varname(varname, other_trade_cred)


    def _calculate_field_amt_due_to_dir(self):
        answer = self.user_inputs.at["amount_due_to_director", "Answer"]

        varname = "current_liab_amount_due_to_director"
        
        if (answer == "NA") or (pd.isnull(answer)):
            pass
        else: 
            client_acc_list = [i.strip() for i in answer.split(",")]
            amt_due_to_dir = self.tb_main[self.tb_main["Account No"].isin(client_acc_list)]
            amt_due_to_dir = amt_due_to_dir["Value"].sum()

            self.add_bal_to_template_by_varname(varname, amt_due_to_dir)


    def _calculate_field_loans_from_relatedco(self):

        answer = self.user_inputs.at["loans_from_related_co", "Answer"]
        varname = "current_liab_loans_from_related_co"
        
        if (answer == "NA") or (pd.isnull(answer)):
            pass
        else: 
            client_acc_list = [i.strip() for i in answer.split(",")]
            loans_from_relatedco = self.tb_main[self.tb_main["Account No"].isin(client_acc_list)]
            loans_from_relatedco = loans_from_relatedco["Value"].sum()

            self.add_bal_to_template_by_varname(varname, loans_from_relatedco)


    def _calculate_field_other_current_liab(self):

        # Minus amount due to director, loan from related co, and trade creditors to get other current liability
        varname_list = ["current_liab_amount_due_to_director", 
                        "current_liab_loans_from_related_co", 
                        "current_liab_trade_cred_other_other", 
                        "current_liab_trade_cred_fund_mgmt"]
        
        rows = self.mapper_class.varname_to_index.loc[varname_list]
        liab_amount = self.outputdf.loc[rows, "Balance"].sum()

        others_row = self.mapper_class.varname_to_index.at["current_liab_other"]
        self.outputdf.loc[others_row, "Balance"] -= liab_amount


    def _calculate_field_amt_due_from_dir_secured(self):
        
        answer = self.user_inputs.at["amount_due_from_director_secured", "Answer"]
        varname = "current_asset_amount_due_from_director_secured"
        
        if (answer == "NA") or (pd.isnull(answer)):
            pass
        else: 
            client_acc_list = [i.strip() for i in answer.split(",")]
            amt_due_from_dir_sec = self.tb_main[self.tb_main["Account No"].isin(client_acc_list)]
            amt_due_from_dir_sec = amt_due_from_dir_sec["Value"].sum()

            self.add_bal_to_template_by_varname(varname, amt_due_from_dir_sec)


    def _calculate_field_amt_due_from_dir_unsecured(self):
        
        answer = self.user_inputs.at["amount_due_from_director_unsecured", "Answer"]
        varname = "current_asset_amount_due_from_director_unsecured"
        
        if (answer == "NA") or (pd.isnull(answer)):
            pass
        else: 
            client_acc_list = [i.strip() for i in answer.split(",")]
            amt_due_from_dir_unsec = self.tb_main[self.tb_main["Account No"].isin(client_acc_list)]
            amt_due_from_dir_unsec = amt_due_from_dir_unsec["Value"].sum()

            self.add_bal_to_template_by_varname(varname, amt_due_from_dir_unsec)


    def _calculate_field_loan_to_relatedco(self):

        answer = self.user_inputs.at["loans_to_related_co", "Answer"]
        varname = "current_asset_loans_to_related_co"
        
        if (answer == "NA") or (pd.isnull(answer)):
            pass
        else: 
            client_acc_list = [i.strip() for i in answer.split(",")]
            loan_to_relatedco = self.tb_main[self.tb_main["Account No"].isin(client_acc_list)]
            loan_to_relatedco = loan_to_relatedco["Value"].sum()

            self.add_bal_to_template_by_varname(varname, loan_to_relatedco)


    def collect_manual_inputs(self):
        '''
        To run only when the user_inputs parameter is not specified when
        the class is initialised.
        '''
        question_list = [
            "List of client names related to fund management (trade debtors): ",
            "Total trade creditors amount: $", 
            "Trade creditors for fund managment amount: $",
            "Enter the client account numbers for amounts due to director or connected persons: ",
            "Enter the client account numbers for loans from related company or associated persons: ",
            "Enter the client account numbers for amounts due from director and connected persons (secured): ",
            "Enter the client account numbers for amounts due from director and connected persons (unsecured): ",
            "Enter the client account numbers for loans to related company or associated person: "]
        
        varname_list = ['trade_debt_fund_mgmt', 
                        'total_trade_cred', 
                        'trade_cred_fund_mgmt', 
                        'amount_due_to_director', 
                        'loans_from_related_co', 'amount_due_from_director_secured', 
                        'amount_due_from_director_unsecured', 
                        'loans_to_related_co']

        self.user_inputs = pd.DataFrame({"Question": question_list,
                                       "Answer": ""}, index=varname_list)

        self.user_inputs["Answer"] = self.user_inputs["Question"].apply(input)
        

    def abs_of_balance_column(self):
        """
        Take absolute of Balance column in self.outputdf 

        Except for unappropriated profit or loss > take the -ve of the number in this row
        """
        # Take the negative of unappropriated profit or loss (because through recalculation from TB, profit is -ve and loss is +ve but should be presented in the form as profit +ve and loss -ve)

        varname = "puc_unappr_profit_or_loss"
        row = self.mapper_class.varname_to_index.at[varname]
        self.outputdf.at[row, "Balance"] = -self.outputdf.at[row, "Balance"]

        # Take absolute of all other rows
        rows = self.outputdf[self.outputdf.index != row].index
        self.outputdf.loc[rows, "Balance"] = self.outputdf.loc[rows, "Balance"].abs()


    def column_mapper(self):

        # Map the Balance amounts to the correct field in F1; whether in the Amount or Subtotal column
        for i in self.outputdf.index:
            if pd.notna(self.outputdf.at[i,"Amount"]) and pd.notna(self.outputdf.at[i+1,"Amount"]):
                if pd.isna(self.outputdf.at[i,"Balance"]):
                    subtotal = subtotal
                else:
                    subtotal += self.outputdf.at[i,"Balance"]
            elif pd.notna(self.outputdf.at[i,"Amount"]) and pd.isna(self.outputdf.at[i+1,
                    "Amount"]):
                if pd.isna(self.outputdf.at[i,"Balance"]) and subtotal != 0:
                    self.outputdf.at[i,"Subtotal"] = subtotal
                    subtotal = 0
                elif pd.notna(self.outputdf.at[i,"Balance"]):
                    subtotal += self.outputdf.at[i,"Balance"]
                    self.outputdf.at[i,"Subtotal"] = subtotal
                    subtotal = 0
            else: 
                subtotal = 0

        for i in self.outputdf.index:
            if pd.notna(self.outputdf.at[i, "Amount"]):
                self.outputdf.at[i, "Amount"] = self.outputdf.at[i, "Balance"]
            elif pd.isna(self.outputdf.at[i, "Amount"]) and pd.notna(self.outputdf.at[i,
                    "Subtotal"]):
                self.outputdf.at[i, "Subtotal"] = self.outputdf.at[i, "Balance"]

    def get_row_totals(self):
        
        # Total Shareholders' Funds or Net Head Office Funds

        varname = "total_shareholder_fund"

        start_varname = "puc_ord_shares"
        end_varname = "puc_net_head_office_funds"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total trade creditors

        varname = "total_trade_cred"

        start_varname = "current_liab_trade_cred_cis_customer_margin_acct"
        end_varname = "current_liab_trade_cred_other_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Current liabilities

        varname = "total_current_liab"

        start_varname = "total_trade_cred"
        end_varname = "current_liab_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Non-current liabilities

        varname = "total_noncurrent_liab"

        start_varname = "noncurrent_liab_cis_cost"
        end_varname = "noncurrent_liab_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Liabilities

        varname = "total_liab"

        varname_list = ["total_current_liab", "total_noncurrent_liab"]
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[varname_list]

        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Shareholders' Funds or Net Head Office Funds and Liabilities

        varname = "total_shareholder_fund_and_liab"

        varname_list = ["total_shareholder_fund", "total_liab"]
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[varname_list]

        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total trade debtors

        varname = "total_trade_debt"

        start_varname = "current_asset_trade_debt_cis_customer_margin_acct"
        end_varname = "current_asset_trade_debt_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")

        
        # Net trade debtors

        varname = "net_trade_debt"

        row = self.mapper_class.varname_to_index.at["total_trade_debt"]
        add = self.outputdf.loc[row, "Subtotal"].sum()

        start_varname = "current_asset_trade_debt_provision_contingency"
        end_varname = "current_asset_trade_debt_provision_bad_debt"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        less = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        total = add-less

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Current asset

        varname = "total_current_asset"

        start_varname = "net_trade_debt"
        end_varname = "current_asset_other_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")


        # Total Non-current asset

        varname = "total_noncurrent_asset"

        start_varname = "noncurrent_asset_fixed_asset"
        end_varname = "noncurrent_asset_other"
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[start_varname:end_varname]
        
        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")

        
        # Total asset
        
        varname = "total_asset"

        varname_list = ["total_current_asset", "total_noncurrent_asset"]
        rows_to_sum = \
            self.mapper_class.varname_to_index.loc[varname_list]

        total = self.outputdf.loc[rows_to_sum, "Subtotal"].sum()

        self.add_bal_to_template_by_varname(varname, total, "y")

    def load_output_to_lunahub(self):

        self.client_number = self.client_class.retrieve_client_info("CLIENTNUMBER")

        loader_class = lunahub.tables.fs_masf1_output.MASForm1Output_LoaderToLunaHub(self.outputdf,
                                                                                     self.client_number,
                                                                                     self.fy
                                                                                     )
        loader_class.main()

    def process_ocr_output(self):

        ocr_df = self.ocr_class.execute()

        column_mapper = {"var_name" : "var_name",
                         "amount"   : "Amount",
                         "subtotal" : "Subtotal"}
        
        ocr_df = ocr_df[column_mapper.keys()]
        # Map col names
        ocr_df = ocr_df.rename(columns = column_mapper)

        self.ocr_df = ocr_df
        
        return ocr_df



if __name__ == "__main__":
        
    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    print (f"Your luna library is at {luna_folderpath}.")
    
    # Get the template folderpath
    template_folderpath = os.path.join(luna_folderpath, "templates")
    
    # AGED RECEIVABLES
    if False:
        # aged_receivables_fp = os.path.join(template_folderpath, "aged_receivables.xlsx")
        aged_receivables_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\Form 1\f1 input data\clean_AR_listing.xlsx"
        print (f"Your aged_receivables_fp is at {aged_receivables_fp}.")
        
        # Load the AR class
        aged_ar_class = common.AgedReceivablesReader_Format1(aged_receivables_fp, 
                                                        sheet_name = 0,            # Set the sheet name
                                                        variance_threshold = 0.1) # 1E-9) # To relax criteria if required.
        aged_ar_class.main()
        
        aged_group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                           ">90": ["91 - 120", "121 - 150", "150+"]}
        
        # Then we get the AR by company (index) and by new bins (columns)
        aged_df_by_company = aged_ar_class.get_AR_by_new_groups(aged_group_dict)
        
    # TB
    if False:
        # tb_fp = os.path.join(template_folderpath, "tb.xlsx")
        mg = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\Form 1\f1 input data\Myer Gold Investment Management - 2022 TB.xlsx"
        ci = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\TB with updated LS codes\Crossinvest TB reclassed.xlsx"
        icm = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\TB with updated LS codes\icm TB reformatted.xlsx"

        tb_fp = mg
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

    if True:
    # Read user_inputs dataframe
        df0 = pyeasylib.excellib.read_excel_with_xl_rows_cols (mas_tb_mapping_fp, sheet_name="Form 1 - Qns for user inputs")

        # Get main table using required header and set index
        REQUIRED_HEADERS = ["Index", "Question", "Answer"]
        user_inputs = pyeasylib.pdlib.get_main_table_from_df(df0, REQUIRED_HEADERS)
        user_inputs.set_index("Index", inplace=True)

        # Fill answer column > depending on company, replace the variable for user_inputs["Answer"]
        mg_answer_list = ["Harvest Platinium International Limited, Equity Summit Limited, Albatross Group, Nido Holdings Limited, Albatross Platinium VCC, Teo Joo Kim or Gerald Teo Tse Sian Or Teo, Oyster Enterprises Limited, Oyster Enterprises Limited, Lawrence Barki, Nico Gold Investments Ltd, UNO Capital Holdings Inc, Boulevard Worldwide Limited, Apollo Pte Limited, CAMSWARD PTE LTD, Granada Twin Investments, UNO Capital Holdings Inc, T & T Strategic Limited, Myer Gold Allocation Fund, Nasor International Limited, Tricor Services (BVI) Limited, Penny Yap, White Lotus Holdings Limited", 
                                   "0", "0", "2-2310, 2-2312", "NA", 
                                   "1-2420, 1-2452", "NA", "1-2448, 1-2450"]
        ci_answer_list = ["NA", "281060", "0", "NA", "NA", "NA", "NA", "NA"]
        icm_answer_list = ["NA", "2902", "0", "NA", "NA", "NA", "NA", "NA"]

        user_inputs["Answer"] = mg_answer_list


    if True:
    # CLASS
        

        ## MG
        client_number = 7167
        client_name = "MYER GOLD INVESTMENT MANAGEMENT PTE. LTD."
        fy_end_date = pd.to_datetime('31 Dec 2022')
        fy = fy_end_date.year
       
        
        # Load tb class from LunaHub
        tb_class = common.TBLoader_From_LunaHub(client_number, fy)
        
        # Load aged ar class
        aged_ar_class = common.AgedReceivablesLoader_From_LunaHub(client_number, fy)

        # Load client class from LunaHub
        client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number)

        # ocr class
        ocr_fn = f"mas_form1_{client_number}_{fy}_alteryx_ocr.xlsx"
        ocr_fp = os.path.join(luna_folderpath, "personal_workspace", "tmp", ocr_fn)
        ocr_class = fsvi.mas.form1.mas_f1_ocr_output_formatter.OCROutputProcessor(filepath = ocr_fp, sheet_name = "Sheet1", form = "form1", luna_fp = luna_folderpath)

        # load user response
        user_response_class = lunahub.tables.fs_masf1_userresponse.MASForm1UserResponse_DownloaderFromLunaHub(
            client_number,
            fy)
        #user_inputs = user_response_class.main()
        
        self = MASForm1_Generator(tb_class, aged_ar_class,
                                mapper_class, client_class, ocr_class,
                                fy=fy, fuzzy_match_threshold=80, 
                                user_inputs=user_inputs)
        
        # To test, if user_inputs parameter is not specified, it will collect user inputs interactively
        # self = MASForm1_Generator(tb_class, aged_ar_class,
        #                         mapper_class, fy=fy, fuzzy_match_threshold=80)
        
        # MG Inputs 
        # Harvest Platinium International Limited, Equity Summit Limited, Albatross Group, Nido Holdings Limited, Albatross Platinium VCC, Teo Joo Kim or Gerald Teo Tse Sian Or Teo, Oyster Enterprises Limited, Oyster Enterprises Limited, Lawrence Barki, Nico Gold Investments Ltd, UNO Capital Holdings Inc, Boulevard Worldwide Limited, Apollo Pte Limited, CAMSWARD PTE LTD, Granada Twin Investments, UNO Capital Holdings Inc, T & T Strategic Limited, Myer Gold Allocation Fund, Nasor International Limited, Tricor Services (BVI) Limited, Penny Yap, White Lotus Holdings Limited
        # 0 
        # 0
        # 2-2310, 2-2312
        # NA
        # 1-2420, 1-2452
        # NA
        # 1-2448, 1-2450

        # CI Inputs
        # NA
        # 281060
        # 0
        # NA
        # NA
        # NA
        # NA
        # NA

        # ICM Inputs
        # need to change fy to 2023
        # NA
        # 2902
        # 0
        # NA
        # NA
        # NA
        # NA
        # NA

        # self.outputdf.to_excel("f1_map_icm.xlsx")
