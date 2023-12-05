'''
Sample script for Dacia and Jia Wey
'''

# Import standard libs
import os
import datetime

# Import luna package and fsvi package

import luna
import luna.common as common
import luna.fsvi as fsvi
from luna.fsvi.mas.template_reader import MASTemplateReader_Form3

import pandas as pd

class MASForm3_Generator:
    
    def __init__(self, 
                 tb_class, mapper_class,
                 fy = 2022):
        
        
        self.tb_class       = tb_class
        self.mapper_class   = mapper_class
        self.fy = fy
        
        self.main()

        
    def main(self):
        
        self._map_varname_to_lscodes()

        self._prepare_output_container()

        # Map fields 
        self.output_debt_accounts()
        self.collect_manual_inputs()
        self.map_debts()
        self.mapSalary()
        self.mapDirector()           
        self.mapCommrebates()
        self.mapReits_baseperf()
        self.mapReits_trans()
        
        self.mapMgmtfees()
        self.mapAdvfees()
        self.map_Corpfinfees_IPO()
        self.map_Corpfinfees_others()
        self.map_trust_custdn()
        self.mapIntrev()
        self.mapDiv()        

        self.mapIntexp()
        self.mapTax()


        self.mapOtherrev()
        self.mapTotalrev()
        self.mapTotalexp()
        self.mapOtherexp()
        self.mapNetprofit()

    def _prepare_output_container(self):

      	# make a copy of the input template
        outputdf = self.mapper_class.df_processed.copy()

        # Add an output column
        outputdf["Balance"] = None

        # Save as attr
        self.outputdf = outputdf

    def add_bal_to_template_by_varname(self, varname, value):
        
        row = self.mapper_class.varname_to_index.at[varname]

        # Map value to template's Balance column
        self.outputdf.loc[row, "Balance"] = value
                
        # Excel row 
        print(f"The varname ({varname}) is at ExcelRow: {row}")

    def _map_varname_to_lscodes(self):
        
        mapper_class = self.mapper_class
        tb_class     = self.tb_class
        
        # get varname to ls code from mapper
        varname_to_lscodes = mapper_class.varname_to_lscodes
        
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

        #
        self.tb_columns_main = tb_columns
        self.tb_with_varname = tb_df.copy()

    def filter_tb_by_varname(self, varname):
        
        tb = self.tb_with_varname.copy()
        
        filtered_tb = tb[tb[varname]][self.tb_columns_main]
        
        return filtered_tb

    # User inputs
    def output_debt_accounts(self):
                
        tb = tb_class.get_data_by_fy(fy)

        baddebt_df = tb[tb["Name"].str.match(fr'(?i)bad\sdebts?')]
        provdebt_df = tb[tb["Name"].str.match(fr'(?i)pro.*\sdebts?|do.+\sdebts?')]
        
        #Save as attr 
        self.baddebt_df = baddebt_df
        self.provdebt_df = provdebt_df

        print(self.baddebt_df)
        print(self.provdebt_df)

    def collect_manual_inputs(self):
        question_list = ["Enter the Account No for bad debts (NA if none): ",  
                         "Enter the Account No for provision debts (NA if none): ", 
                         "Enter Director's remuneration (0.00 if None): $ "]
        
        self.inputs_df = pd.DataFrame({"Questions": question_list, 
                                       "Answers": ""})
        
        self.inputs_df["Answers"] = self.inputs_df["Questions"].apply(input)
        
    def map_debts(self):
        varname1 = "exp_bad_debts"
        varname2 = "exp_prov_dtf_debts"

        client_baddebt_list = [i.strip() for i in self.inputs_df.at[0, "Answers"].split(",")]
        client_provdebt_list= [i.strip() for i in self.inputs_df.at[1, "Answers"].split(",")]

        if (client_baddebt_list == "NA") & (client_provdebt_list == "NA"):
            print("Note: No bad debts accounts found. The amount will be set to 0.00. ")
            bad_debt = 0.00
            provision_debt = 0.00

        else:
            # Amt 
            bad_debt = self.baddebt_df[self.baddebt_df["Account No"].isin(client_baddebt_list)]["Value"].sum()
            provision_debt = self.provdebt_df[self.provdebt_df["Account No"].isin(client_provdebt_list)]["Value"].sum()
            
        # Map amt to Template
        self.add_bal_to_template_by_varname(varname1, bad_debt)
        self.add_bal_to_template_by_varname(varname2, provision_debt)

        # to delete
        print(f" Mapped {varname1} : ${bad_debt}")      
        print(f" Mapped {varname2} : ${provision_debt}")      

    def mapDirector(self):
        varname = "exp_dir_renum"

        # Amt
        director = float(self.inputs_df.at[2,"Answers"])

        # Map amt to Template 
        self.add_bal_to_template_by_varname(varname, director)

        # to delete
        print(f" Mapped {varname} : {director}")    

        return director

    def mapSalary(self):

        director = self.mapDirector()

        varname = "exp_slry"

        # Get 7300 series 
        interval_list = [pd.Interval(7300,7400, closed='left')]
        
        boolean, true_match, false_match = \
            self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, interval_list)
        
        # Amt 
        total_salary = true_match["Value"].sum()
        print(total_salary)
        salary = total_salary - director   

        # Map amt to Template
        self.add_bal_to_template_by_varname(varname, salary)

        # to delete
        print(f" Mapped {varname} : {salary}")    


    def mapCommrebates(self):
        varname = "rev_comm_rebates"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            commission_rebate = 0.00 
            print(f"No commission rebate accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else:
            commission_rebate = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, commission_rebate)  
            
        # to delete
        print(f" Mapped {varname} : {commission_rebate}")    
    
    def mapReits_baseperf(self):
        varname = "rev_reit_mgtfees_baseperf"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            base_perf = 0.00
            print(f"No Reits (base performance) accounts. The amount will be set to 0.00 {filtered_tb}")

        else:
            base_perf = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, base_perf)  

        # to delete
        print(f" Mapped {varname} : {base_perf}")    

    def mapReits_trans(self):
        varname = "rev_reit_mgtfees_trans"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)
        
        if filtered_tb.empty:
            trans = 0.00
            print(f"No REITs (transaction fees) accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else: 
            trans = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, trans)  

        # to delete
        print(f" Mapped {varname} : {trans}")    

    def mapMgmtfees(self):
        varname = "rev_pfl_mgtfees_mgmt_fees"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            mgmt_fees = 0.00
            print(f"No Management fees accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else: 
            mgmt_fees = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, mgmt_fees)  

        # to delete
        print(f" Mapped {varname} : {mgmt_fees}")    

    def mapAdvfees(self):
        varname = "rev_pfl_mgtfees_adv_fees"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            advisory_fees = 0.00
            print(f"No Advisory fees accounts. The amount will be set to 0.00 {filtered_tb}")

        else:  
            advisory_fees = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, advisory_fees)

        # to delete
        print(f" Mapped {varname} : {advisory_fees}")    

    def map_Corpfinfees_IPO(self):
        varname = "rev_corp_finfees_ipo"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)
        
        if filtered_tb.empty:
            corpfinfees_IPO = 0.00
            print(f"No Corporate finance (IPO) fees accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else: 
            corpfinfees_IPO = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, corpfinfees_IPO)  

        # to delete
        print(f" Mapped {varname} : {corpfinfees_IPO}")    
    
    def map_Corpfinfees_others(self):
        varname = "rev_corp_finfees_others"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            corpfinfees_others = 0.00
            print(f"No Corporate finance (Others) fees accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else:
            corpfinfees_others = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, corpfinfees_others)      

        # to delete
        print(f" Mapped {varname} : {corpfinfees_others}")    

    def map_trust_custdn(self):
        varname = "rev_trust_custdn"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)
        
        if filtered_tb.empty:
            trust_custdn = 0.00
            print(f"No Trustee and custodian fees accounts. The amount will be set to 0.00 {filtered_tb}")

        else:       
            trust_custdn = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, trust_custdn)   
            
        # to delete
        print(f" Mapped {varname} : {trust_custdn}")    
   
    def mapIntrev(self):
        varname = "rev_int_others"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            rev_int_others = 0.00
            print(f"No Interest (Others) accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else:
            rev_int_others = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, rev_int_others)

        # to delete 
        print(f" Mapped {varname} : {rev_int_others}")    

    def mapDiv(self):
        varname = "rev_dividend"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            dividend = 0.00
            print(f"No Dividend accounts. The amount will be set to 0.00 {filtered_tb}")
        
        else: 
            dividend = abs(filtered_tb["Value"].sum())

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, dividend)

        # to delete 
        print(f" Mapped {varname} : {dividend}")    


    def mapIntexp(self):
        varname = "exp_int_expense"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            exp_int_expense = 0.00
            print(f"No Interest expense accounts. The amount will be set to 0.00 {filtered_tb}")

        else:
            exp_int_expense = filtered_tb["Value"].sum()

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, exp_int_expense)

        # to delete 
        print(f" Mapped {varname} : {exp_int_expense}")    

    def mapTax(self):
        varname = "tax"

        # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            corp_tax = 0.00
            print(f"No tax accounts. The amount will be set to 0.00 {filtered_tb}")

        else: 
            corp_tax = filtered_tb["Value"].sum()

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, corp_tax)
                   
        # to delete
        print(f" Mapped {varname} : {corp_tax}")   

    def mapOtherrev(self):
        varname = "rev_other_revenue"

         # Get df by varname
        filtered_tb = self.filter_tb_by_varname(varname)

        if filtered_tb.empty:
            other_rev = 0.00
            print(f"No Other revenue accounts. The amount will be set to 0.00 {filtered_tb}")

        else: 
            interval_list = [pd.Interval(7410.400, 7410.400, closed = "both")]
            boolean, true_match, false_match =\
                  self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, interval_list)  
            
            # sum L/S 7410.4
            forex = true_match["Value"].sum() 

            # Sum negative balances 
            other_rev = filtered_tb[~filtered_tb["L/S"].isin(['7410.4'])]
            other_rev = other_rev[other_rev["Value"] <0]["Value"].sum()


            # if forex gain, add to other_rev
            if forex < 0:
                
                other_rev = abs(other_rev) + abs(forex)

            else: 
                other_rev = abs(other_rev)


        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, other_rev)  

        # to delete
        print(f" Mapped {varname} : {other_rev}")  
  

    def mapTotalrev(self):

    # Get other rev from Template
        otherrev_varname = "rev_other_revenue"
        otherrev_row = self.mapper_class.varname_to_index.at[otherrev_varname]
        other_rev = self.outputdf.at[otherrev_row,"Balance"]

    # First method: Compute from Template 
        varname = "rev_total_revenue"

        # Get index row
        row = self.mapper_class.varname_to_index.at[varname]

        # Amt
        total_revenue = abs(self.outputdf.loc[:row, "Balance"].sum())
        print(f"total revenue computed from Template: {total_revenue}")

    # Second method: Compute from TB's Value
        
        # All revenue accounts without L/S 7410.400 (net forex), L/S 7410.100 (other rev), L/S 7410.200 (forex gain), L/S 7410.300(forex loss)
        exclude_ls = [pd.Interval(5000,6900.4, closed = "both"), 
                      pd.Interval(7410.100,7410.400, closed='both')]
        
        boolean, true_match, false_match = \
            self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, exclude_ls)
        
        rev = false_match[false_match["Class"].str.match(r'Revenue.+')]

        rev = rev["Value"].sum()
        
        print(f"rev is {rev}")

        total_rev = abs(rev) + abs(other_rev)
        
        if total_revenue == total_rev: 
            print("Total Revenue Match", total_revenue, total_rev)
        else: 
            print("Total Revenue does not Match", total_revenue, total_rev)

        # Map amt to Template  
        self.add_bal_to_template_by_varname(varname, total_rev)  
          
    def mapTotalexp(self):
        
    # Compute from TB's Value 
        
        # Accounts without L/S 7410.400 (net forex), L/S 7410.100 (other rev),
        # L/S 7410.200 (forex gain), L/S 7410.300(forex loss), L/S 7500 (tax)

        exclude_ls = [pd.Interval(5000.000, 6900.400, closed = 'both'), 
                      pd.Interval(7410.100,7410.400, closed='both'), 
                      pd.Interval(7500.000, 7500.000, closed = "both")]
        
        boolean, true_match, false_match = \
            self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, exclude_ls)
        
        # Filter for Expense accounts and sum amt
        total_exp = false_match[false_match["Class"].str.match(r'Expense.+')]
        total_exp = total_exp["Value"].sum()

        # Sum L/S 7410.1 thar are positive balances
        varname = "rev_other_revenue"
        filtered_tb = self.filter_tb_by_varname(varname)
        other_rev_pos = filtered_tb[filtered_tb["L/S"].isin(["7410.1"])]
        other_rev_pos = other_rev_pos[other_rev_pos["Value"] > 0]["Value"].sum()

        # Update total expense amt 
        total_exp = total_exp + other_rev_pos
        
        # Filter if L/S code is '7410.200' or '7410.300' or '7410.400'
        # if net forex is <0 then we add this to total expense
        interval_list = [pd.Interval(7410.200, 7410.300, closed = "both")]
        boolean, true_match, false_match = self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, interval_list)    
        
        if not true_match.empty:

            net_forex = true_match["Value"].sum()

            if net_forex > 0: 
                print("Forex loss", net_forex)
                total_exp = total_exp + net_forex

            else: 
                print("Forex gain", net_forex)
                total_exp = total_exp

        else:

            interval_list2 = [pd.Interval(7410.400, 7410.400, closed = "both")]
            boolean, true_match, false_match = self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, interval_list2)    
        
            net_forex = true_match["Value"].sum()

            if net_forex > 0: 
                print("Forex loss", net_forex)
                total_exp = total_exp + net_forex

            else: 
                print("Forex gain", net_forex)
                total_exp = total_exp
        
        # Map amt to Template  
        varname = "exp_total_expense"
        self.add_bal_to_template_by_varname(varname, total_exp)

        print(f"total exp is {total_exp}")

    def mapOtherexp(self):

        # Get Total exp amt from Template 
        totalexp_varname = "exp_total_expense"
        totalexp_row = self.mapper_class.varname_to_index.at[totalexp_varname]
        print(type(totalexp_row))  
        total_exp = self.outputdf.at[totalexp_row,"Balance"]
        print(type(total_exp))  

        # Get index row 
        salary_varname = "exp_slry" 
        debt_varname = "exp_bad_debts"
 
        debt_row = self.mapper_class.varname_to_index.at[debt_varname] 
        print(type(debt_row))         
        salary_row = self.mapper_class.varname_to_index.at[salary_varname]
        print(type(salary_row))  
        
        # Amt (Total exp - sum of balances from (2a) Debts to (2i) Salaries)
        other_exp = total_exp - self.outputdf.loc[debt_row:salary_row, "Balance"].sum()
        print(type(other_exp))  
        
        # Map amt to Template 
        varname = "exp_other_expense"
        self.add_bal_to_template_by_varname(varname, other_exp)  

        # to delete
        print(f" Mapped {varname} : {other_exp}")  

    def mapNetprofit(self):

    # Net profit before tax
        # First method: Compute from TB's (all 7000++ accounts except tax)
        exclude_ls = [pd.Interval(7000.000,7500.000, closed = 'left')]

        boolean, true_match, false_match = \
            self.tb_class.filter_tb_by_fy_and_ls_codes(self.fy, exclude_ls)
        
        net_profit = true_match["Value"].sum() *-1

        # Second Method: Compute from Template
        # Get total rev & total exp from Template
        rev_varname = "rev_total_revenue"
        total_rev_row = self.mapper_class.varname_to_index.at[rev_varname]
        total_rev = self.outputdf.at[total_rev_row,"Balance"]

        exp_varname = "exp_total_expense"
        total_exp_row = self.mapper_class.varname_to_index.at[exp_varname]
        total_exp = self.outputdf.at[total_exp_row,"Balance"]
        net_profit_2 = total_rev - total_exp

        if net_profit_2 == net_profit: 
            print("Net profit before tax Match", net_profit, net_profit_2)
        else: 
            print("Net profit before tax does not Match", net_profit, net_profit_2)

        # Map amt to Template 
        npbt_varname = "npbt"
        self.add_bal_to_template_by_varname(npbt_varname, net_profit)  

    # Net profit after tax but before extraordinary items
        # Get tax amt from Template
        tax_varname = "tax"
        tax_row = self.mapper_class.varname_to_index.at[tax_varname]
        tax = self.outputdf.loc[tax_row,"Balance"]

        npbt_row = self.mapper_class.varname_to_index.at[npbt_varname]    
        npat_bef_extraord = self.outputdf.loc[npbt_row, "Balance"] - tax
        npat_bef_varname = "npat_bef_extraord"
        self.add_bal_to_template_by_varname(npat_bef_varname, npat_bef_extraord)      

        # Net profit after tax and extraordinary items for the year
        npat_varname = "npat"
        npat = self.outputdf.loc[npbt_row+2, "Balance"] + 0.00
        self.add_bal_to_template_by_varname(npat_varname, npat)   

        # to delete
        print(f" Mapped {npbt_varname} : {net_profit}") 
        print(npat_bef_extraord)
        print(f" Mapped {npat_varname} : {npat}") 


if __name__ == "__main__":
        
    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)
    print (f"Your luna library is at {luna_folderpath}.")
    
    # Get the template folderpath
    template_folderpath = os.path.join(luna_folderpath, "templates")
    
    # AGED RECEIVABLES
    if False:
        aged_receivables_fp = os.path.join(template_folderpath, "aged_receivables.xlsx")
        aged_receivables_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\aged_receivables_template.xlsx"
        print (f"Your aged_receivables_fp is at {aged_receivables_fp}.")
        
        # Load the AR class
        aged_ar_class = common.AgedReceivablesReader_Format1(aged_receivables_fp, 
                                                        sheet_name = 0,            # Set the sheet name
                                                        variance_threshold = 0.1) # 1E-9) # To relax criteria if required.
        
        aged_group_dict = {"0-90": ["0 - 30", "31 - 60", "61 - 90"],
                           ">90": ["91 - 120", "121 - 150", "150+"]}
        
        # Then we get the AR by company (index) and by new bins (columns)
        aged_df_by_company = aged_ar_class.get_AR_by_new_groups(aged_group_dict)
        
    # TB
    if True:
        #tb_fp = os.path.join(template_folderpath, "tb.xlsx")
        #tb_fp = r"D:\Desktop\owgs\CODES\luna\personal_workspace\dacia\Myer Gold Investment Management - 2022 TB.xlsx"
        
        # MG
        #tb_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\TB with updated LS codes\Myer Gold Investment Management - 2022 TB.xlsx"
        
        # CrossInvest
        tb_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\TB with updated LS codes\Crossinvest TB reclassed.xlsx"
        
        # ICM
        #tb_fp = r"P:\YEAR 2023\TECHNOLOGY\Technology users\FS Vertical\TB with updated LS codes\icm TB reformatted.xlsx"
        
        
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
    if False:
        
        mas_tb_mapping_fp = os.path.join(luna_folderpath, "parameters", "mas_forms_tb_mapping.xlsx")
        print (f"Your mas_tb_mapping_fp is at {mas_tb_mapping_fp}.")
        
        # Load the class
        mapper_class = fsvi.mas.MASTemplateReader_Form1(mas_tb_mapping_fp, sheet_name = "Form 1 - TB mapping")
    
        # process df is here:
        df_processed = mapper_class.df_processed  # need to build methods

    # Form 3 mapping 
    if True:
        
        mas_tb_mapping_fp = os.path.join(luna_folderpath, "parameters", "mas_forms_tb_mapping.xlsx")
        print (f"Your mas_tb_mapping_fp is at {mas_tb_mapping_fp}.")
        
        # Load the class
        mapper_class = MASTemplateReader_Form3(mas_tb_mapping_fp, sheet_name = "Form 3 - TB mapping")
    
        # process df is here:
        df_processed = mapper_class.df_processed  # need to build methods
        


    # CLASS
    fy=2022
    self = MASForm3_Generator(tb_class,
                              mapper_class, fy=fy) # is this an instance of a class?
    
    
     # Get previous balance  
    prevfy = 2021
    temp = MASForm3_Generator(tb_class,
                              mapper_class, fy=prevfy)
    self.outputdf['Previous Balance'] = temp.outputdf["Balance"]

    # Reorder columns by index 
    column_order = [0,1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 10] 
    self.outputdf = self.outputdf.iloc[:, column_order]  

    # Output to excel 
    self.outputdf.to_excel("draft.xlsx") 



    # MG
    # NA, NA, 748216
    # NA, NA, 792856

    # CrossInvest
    # 1590054, CW9, 677754
    # 1590054, CW9, 819860

    # ICM
    # change fy to 2023
    # NA, NA, 619423
    