import pandas as pd
import numpy as np
import re
from fuzzywuzzy import fuzz, process
from datetime import datetime
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation


class TBReader:
    def __init__(self, file, fy):
        # Initialize file and fy
        self.file = file
        self.fy = fy
        self.df = pd.read_excel(self.file)

    def tbprocessing(self):
        df = self.df
        df.fillna("NA", inplace=True)

        # Rename year columns
        yr = df.columns[4:].to_series().dt.year
        yr = pd.Series(data=yr.index, index=yr.values)
        col_names = [str(i).strip() for i in df.columns[:4]]
        col_names = col_names + yr.index.tolist()
        df.columns = col_names

        # Melt dataframe to reshape it
        df = df.melt(id_vars=['Account No', 'Name', 'L/S', 'Class'],
                     value_vars=yr.index, var_name="FY", value_name="Amount")

        df['Account No'] = df['Account No'].astype(str)
        df = df.sort_values(['Account No', 'FY'])
        df = df[df["FY"] == self.fy]
        df.rename(columns = lambda x: str(x).strip(), inplace = True)
        return df

    # ICM TB is in different format than CI and MG, process_icm to match MG TB format
    def process_icm(self):
        # Extract same columns 
        self.df = self.df.iloc[:,[0,1,6,8,-1]]

        # Rename year column + Account No column to match MG TB format
        date = re.sub("[^\d\/]", "", self.df.columns[-1])
        date = datetime.strptime(date, "%d/%m/%Y")
        self.df.rename(columns = {
            self.df.columns[-1] : date,
            "Account no": "Account No"}, 
            inplace = True)
        
        return self.df

class TemplateReader:
    def __init__(self, file, sheet_name):
        self.file = file
        self.sheet_name = sheet_name
        self.df2 = pd.read_excel(self.file, sheet_name=sheet_name, skiprows = 6)

    def templateprocessing(self):
        # Process template so that all L/S codes are in 1 column and L/S code is in '0000.000' format
        self.df2[self.df2.columns[0]].fillna(self.df2[self.df2.columns[1]], 
                                             inplace=True)
        self.df2.drop(columns = self.df2.columns[1], inplace=True)
        self.df2.columns = ["Field", "Amount", "Subtotal"]
        self.df2["L/S"] = self.df2.Subtotal.fillna(self.df2.Amount)

        self.df2["L/S (num)"] = self.df2["L/S"].apply(lambda x: str(x).replace(" ", "").split(','))

        self.df2["L/S (num)"] = self.df2["L/S (num)"].apply(
            lambda x: ["%.3f" % pd.to_numeric(i, errors='coerce') for i in x])

        for i in self.df2.index:
            if self.df2["L/S (num)"][i][0] == 'nan':
                self.df2["L/S (num)"][i] = self.df2["L/S"][i]

        self.df2["L/S (num)"] = self.df2["L/S (num)"].apply(lambda x: [] if type(x) is not list else x)

        return self.df2
    
    def col_map(self):
        # Map the Balance amounts to the correct field in F1; whether in the Amount or Subtotal column
        for i in df2.index:
            if pd.notna(df2.at[i,"Amount"]) and pd.notna(df2.at[i+1,"Amount"]):
                if pd.isna(df2.at[i,"Balance"]):
                    subtotal = subtotal
                else:
                    subtotal += df2.at[i,"Balance"]
            elif pd.notna(df2.at[i,"Amount"]) and pd.isna(df2.at[i+1,
                    "Amount"]):
                if pd.isna(df2.at[i,"Balance"]) and subtotal != 0:
                    df2.at[i,"Subtotal"] = subtotal
                    subtotal = 0
                elif pd.notna(df2.at[i,"Balance"]):
                    subtotal += df2.at[i,"Balance"]
                    df2.at[i,"Subtotal"] = subtotal
                    subtotal = 0
            else: 
                subtotal = 0

        for i in df2.index:
            if pd.notna(df2.at[i, "Amount"]):
                df2.at[i, "Amount"] = df2.at[i, "Balance"]
            elif pd.isna(df2.at[i, "Amount"]) and pd.notna(df2.at[i,
                    "Subtotal"]):
                df2.at[i, "Subtotal"] = df2.at[i, "Balance"]

        return df2

    def get_index(self, field, start_field = None, end_field = None):
        # Method to retrieve the indexes of each field
        mask = self.df2[self.df2["Field"].str.contains(
            field, case=False, regex=True, na=False)]
        index_list = mask.index

        if start_field is not None: 
            start_index = self.df2[self.df2["Field"].str.contains(start_field, case=False,regex=True,na=False)].index[0]
            index_list = [i for i in index_list if i > start_index]

        if end_field is not None: 
            end_index = self.df2[self.df2["Field"].str.contains(end_field, case=False,regex=True,na=False)].index[0]
            index_list = [i for i in index_list if i < end_index]
        
        index = index_list[0]
        return index

    def sum_totals(self):
        # To get the sum of balance lines

        # Take absolute amount of both columns
        df2[["Amount", "Subtotal"]] = df2[["Amount", "Subtotal"]].abs()

        # Total Shareholders' Funds or Net Head Office Funds
        end_index = self.get_index("total shareh")
        df2.at[end_index, "Subtotal"] = df2.loc[:end_index-1,"Subtotal"].sum()
        shareholder_funds = df2.at[end_index, "Subtotal"]

        # Total trade creditors
        start_index = end_index+1
        end_index = self.get_index("total trade.*cred")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()

        # Total Current Liabilties
        start_index = end_index
        end_index = self.get_index("total cur.*liab")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()
        curr_liab = df2.at[end_index, "Subtotal"]

        # Total Non-Current Liabilities
        start_index = end_index+1
        end_index = self.get_index("total non.*cur.*liab")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()
        non_curr_liab = df2.at[end_index, "Subtotal"]

        # Total Current and Non-Current Liabilities
        end_index = self.get_index("total cur.*and.*liab")
        df2.at[end_index, "Subtotal"] = curr_liab + non_curr_liab
        total_liab = df2.at[end_index, "Subtotal"]

        # Total Shareholders' Funds or Net Head Office Funds and Liabilities
        end_index = self.get_index("total shareh.*and.*liab")
        df2.at[end_index, "Subtotal"] = shareholder_funds + total_liab

        # Total trade debtors
        start_index = end_index+1
        end_index = self.get_index("total trade.*debt")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()

        # Net trade debtors
        start_index = end_index
        end_index = self.get_index("Net trade.*deb")
        less = df2.loc[start_index+1:end_index-1,"Subtotal"].sum()
        df2.at[end_index, "Subtotal"] = df2.at[start_index, "Subtotal"] - less

        # Total Current Assets
        start_index = end_index
        end_index = self.get_index("total cur.*assets")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()
        curr_assets = df2.at[end_index, "Subtotal"]

        # Total Non-Current Assets
        start_index = end_index+1
        end_index = self.get_index("total non.*asset")
        df2.at[end_index, "Subtotal"] = df2.loc[start_index:end_index-1,"Subtotal"].sum()
        non_curr_assets = df2.at[end_index, "Subtotal"]

        # Total Current and Non-Current Assets
        end_index = self.get_index("total curr.*and.*non.*asset")
        df2.at[end_index, "Subtotal"] = curr_assets + non_curr_assets

        return df2

class GeneralAcc:
    def __init__(self, df):
        self.df = df

    def aggregate(self, list):
        # Aggregate amounts for given list of L/S codes
        ga = self.df[self.df["L/S"].isin(list)].groupby("FY").agg(
            {"FY": 'first', "Account No": ",".join, "Name": ",".join, "L/S": "first", "Class": 'first', "Amount": "sum"})
        ga = pd.concat([general_accounts, ga])
        return ga

    def one_acc(self, acc):
        # Aggregate amounts for one given L/S code
        gb = self.df[self.df["L/S"] == acc].groupby("FY").agg(
            {"FY": 'first', "Account No": ",".join, "Name": ",".join, "L/S": "first", "Class": 'first', "Amount": "sum"})
        gb = pd.concat([general_accounts, gb])
        return gb

    def map_acc(self, general_accounts, df2):
        # Map general accounts to df2
        df2["Balance"] = df2["L/S (num)"].apply(
            lambda x: general_accounts[general_accounts["L/S"].isin(x)]["Amount"])
        return df2

class Inputs:
    def __init__(self, df, df2, df3):
        self.df = df
        self.df2 = df2
        self.df3 = df3

    def get_index(self, field, start_field = None, end_field = None):
        mask = self.df2[self.df2["Field"].str.contains(
            field, case=False, regex=True, na=False)]
        index_list = mask.index

        if start_field is not None: 
            start_index = self.df2[self.df2["Field"].str.contains(start_field, case=False,regex=True,na=False)].index[0]
            index_list = [i for i in index_list if i > start_index]

        if end_field is not None: 
            end_index = self.df2[self.df2["Field"].str.contains(end_field, case=False,regex=True,na=False)].index[0]
            index_list = [i for i in index_list if i < end_index]
        
        index = index_list[0]

        return index

    def deposit(self):
        # Calculate deposit and prepayment balances
        dep_prepaid = self.df[self.df["L/S"].isin([5400.1,5200.2])]

        # Create deposit and prepayment indicator
        for i in dep_prepaid.index:
            if re.search("(?i)pre", dep_prepaid.at[i, "Name"]):
                dep_prepaid.at[i, "Indicator"] = "Prepayment"
            elif re.search("(?i)deposit", dep_prepaid.at[i, "Name"]):
                dep_prepaid.at[i, "Indicator"] = "Deposit"
            else:
                dep_prepaid.at[i, "Indicator"] = "Others"

        # Append the amounts to the respective fields
        index = self.get_index("deposit", "other current assets")
        self.df2.at[index, "Balance"] = dep_prepaid[dep_prepaid["Indicator"]
                                                    == "Deposit"].Amount.sum()
        
        index = self.get_index("pre.*pa")
        self.df2.at[index, "Balance"] = dep_prepaid[dep_prepaid["Indicator"]
                                                    =="Prepayment"].Amount.   sum()
        
        index = self.get_index("others", "other current assets", "total current assets")
        amt_mapped = dep_prepaid[(dep_prepaid["L/S"] == 5200.2) & (dep_prepaid["Indicator"] != "Others")].Amount.sum()
        self.df2.at[index, "Balance"] -= amt_mapped

        return self.df2

    def pref_shares(self):
        share_cap = self.df[self.df["L/S"]==6900.1]
    
        for i in share_cap.index:
            # Search for Irredeemable Preference Shares
            if re.search("(?i)irrede.*pref.*", share_cap.at[i, "Name"]):
                # Search for non-cumulative in irredeemable preference shares
                if re.search("(?i)non.*cumu.*", share_cap.at[i, "Name"]):
                    share_cap.at[i, "Tag"] = "Irredeemable Preference Share (Non-Cumulative)"
                # If account name has no tagging of cumulative or non-cumulative, default set as cumulative irredeemable preference share
                else:
                    share_cap.at[i, "Tag"] = "Irredeemable Preference Share (Cumulative)"

            # Search for Redeemable Preference Shares
            elif re.search("(?i)(?<!ir)redeem.*pref.*", share_cap.at[i, "Name"]): 
                # Search for non-current in redeemable preference shares
                if re.search("(?i)non.*current", share_cap.at[i, "Name"]):
                    share_cap.at[i, "Tag"] = "Redeemable Preference Share (Non-Current)"
                # If account name has no tagging of current or non-current, default set as current redeemable preference share
                else: 
                    share_cap.at[i, "Tag"] = "Redeemable Preference Share (Current)"
            # All remaining untagged accounts to tag as ordinary shares
            else:
                share_cap.at[i, "Tag"] = "Ordinary Shares"

        return share_cap

    def map_pref_shares(self, share_cap):
        # Ordinary shares
        index = self.get_index("ord.*share")
        amount = share_cap[share_cap["Name"]=="Ordinary Shares"].Amount.sum()
        df2.at[index, "Balance"] = amount

        # Redeemable Preference Share (Current)
        index = self.get_index("redeem.*pref.*", 
                                end_field="total current liabi")
        amount = share_cap[share_cap["Name"]=="Redeemable Preference Share (Current)"].Amount.sum()
        df2.at[index, "Balance"] = amount
        
        # Redeemable Preference Share (Non-Current)
        index = self.get_index("redeem.*pref", 
                                start_field="non.*liabil")
        amount = share_cap[share_cap["Name"]=="Redeemable Preference Share (Non-Current)"].Amount.sum()
        df2.at[index, "Balance"] = amount

        # Irredeemable Preference Share (Cumulative)
        index = self.get_index("irredeemable and cumu")
        amount = share_cap[share_cap["Name"]=="Irredeemable Preference Share (Cumulative)"].Amount.sum()
        df2.at[index, "Balance"] = amount

        # Irredeemable Preference Share (Non-Cumulative)
        index = self.get_index("irre.*non.*cumu")
        amount = share_cap[share_cap["Name"]=="Irredeemable Preference Share (Non-Cumulative)"].Amount.sum()
        df2.at[index, "Balance"] = amount
        
        return df2

    def verify_share_tagging(self, share_cap):
        print(share_cap)
        a = input("Is the auto tagging correct? (y/n): ").lower()
        if a == 'n':
            print("Please assign correct tags")
            share_cap.to_excel("Share tagging.xlsx")
            wb = openpyxl.load_workbook("Share tagging.xlsx")
            ws = wb.active

            # Create data validation 
            valid_options = '"Ordinary Shares, Irredeemable Preference Share (Cumulative), Irredeemable Preference Share (Non-Cumulative), Redeemable Preference Share (Current), Redeemable Preference Share (Non-Current)"'
            dv = DataValidation(type="list", formula1=valid_options)
            # Error message
            dv.error ='Your entry is not in the list'
            dv.errorTitle = 'Invalid Entry'
            ws.add_data_validation(dv)

            dv.add("H2:H1048576")

            wb.save("Share tagging.xlsx")

            file = input("File path with manual tags: ")

            share_cap = pd.read_excel(file)
        return share_cap

    def trade_creditors(self):
        a = input(
            "Is there any amount in other current liabilities that is related to trade creditors (y/n): ").lower()
        while a != 'n' or a != 'y':
            if a == 'n':
                break
            elif a == 'y':
                # Fuzzy matching with AP listing (currently not in use as no AP listing)
                # totaltradecreditors = float(input("Enter the amount of total trade creditors: "))
                # index = self.get_index("Total trade cred.*")[0]
                # df2.at[index,"Balance"] = totaltradecreditors

                # fundmgmt_df = input("Enter list of supplier names related to fund management: ")
                # fundmgmt_list = [i.strip() for i in fundmgmt_df.split(",")]

                # df4['Match_score'] = df4['Name'].apply(lambda x: process.extractOne(x, fundmgmt_list, scorer=fuzz.token_sort_ratio))

                # # Filter the rows in df4 where Match_score is above threshold of 80
                # df4["Matched?"] = df4["Match_score"].apply(lambda x: x[1]) >= 80

                # # If matched,
                # fundmgmt_creditors = df4[df4["Matched?"] == True].iloc[:,1].sum()

                # index = self.get_index("trade cred.*fund m.*")[0]
                # df2.at[index,"Balance"] = fundmgmt_creditors

                # others = totaltradecreditors - fundmgmt_creditors
                # index_list = self.get_index(field = "other than the above", area = "CURRENT LIABILITIES")
                # start_index = self.get_index(field = "other trade cred.*", area = "CURRENT LIABILITIES")[0]
                # stop_index = self.get_index(field = "total trade cred.*", area = "CURRENT LIABILITIES")
                # index = [i for i in index_list if start_index <= i <= stop_index]
                # df2.at[index[0],"Balance"] = others

                # index = self.get_index(field = "other current liab*")[0]
                # df2.at[index,"Balance"] += totaltradecreditors

                totaltradecreditors = float(
                    input("Enter the amount of total trade creditors: "))
                index = self.get_index("total trade cred")
                df2.at[index, "Balance"] = totaltradecreditors

                fundmgmt = float(
                    input("Enter the amount of trade creditors related to fund management: "))
                index = self.get_index("trade cred.*fund")
                df2.at[index, "Balance"] = fundmgmt

                others = totaltradecreditors-fundmgmt

                index = self.get_index("other than the above", 
                                       start_field="other trade cred", end_field="total trade cred")

                df2.at[index, "Balance"] = others

                index = self.get_index("other current liab")
                df2.at[index, "Balance"] += totaltradecreditors
                break
            else:
                a = input(
                    "Please enter a valid input \nIs there any amount in other current liabilities that is related to trade creditors (y/n): ").lower()
        return df2

    def rpt_l(self):
        a = input("Are there any accounts in other current liabilities that is amount due to director or loans from related company (y/n): ").lower()
        while a != 'n' or a != 'y':
            if a == 'n':
                break
            if a == 'y':
                due_to_dir = input(
                    "Enter the client account numbers for amounts due to director or connected persons (separate account numbers with a comma, enter 0 if not applicable: ")
                due_to_dir = due_to_dir.replace(" ", "").split(",")
                try:
                    index_due_to_dir = self.get_index("due to dire", 
                                                      end_field="total curr.*liab")
                    df2.at[index_due_to_dir, 
                           "Balance"] = df[df["Account No"].isin(due_to_dir)].Amount.sum()

                    index_ocl = self.get_index("other curr.*liab")
                    df2.at[index_ocl,
                           "Balance"] -= df2.at[index_due_to_dir, "Balance"]
                except:
                    print(f"Input received: {due_to_dir}")
                loan = input(
                    "Enter the client account numbers for loans from related company or associated persons (separate account numbers with a comma, enter 0 if not applicable): ")
                loan = loan.replace(" ", "").split(",")
                try:
                    index_loan = self.get_index("loan.*from.*relate", 
                                                end_field="other curr.*liab")
                    df2.at[index_loan, "Balance"] = df[df["Account No"].isin(
                        loan)].Amount.sum()
                    df2.at[index_ocl, "Balance"] -= df2.at[index_loan, "Balance"]
                except:
                    print(f"Input received: {loan}")
                break
            else:
                a = input("Please enter a valid input \nAre there any accounts in other current liabilities that is amount due to director or loans from related company (y/n): ").lower()
        return df2

    def trade_debtors(self):
        a = input(
            "Is there any amount in other trade debtors that is for fund management (y/n): ")
        while a != 'n' or a != 'y':
            if a == 'n':
                break
            elif a == 'y':
                # Fuzzy matching for the names of trade debtors related to fund management
                fundmgmt_df = input(
                    "Enter list of client names related to fund management: ")
                fundmgmt_list = [i.strip() for i in fundmgmt_df.split(",")]

                self.df3['Match_score'] = self.df3['Name'].apply(
                    lambda x: process.extractOne(x, fundmgmt_list, scorer=fuzz.token_sort_ratio))

                # Filter the rows in df3 where 'Match' score is above threshold of 80
                self.df3["Matched?"] = self.df3["Match_score"].apply(
                    lambda x: x[1]) >= 80

                # If matched,
                fundmgmt_debtors = self.df3[self.df3["Matched?"]
                                            == True].iloc[:, 1].sum()

                index = self.get_index("trade debt.*fund m")
                df2.at[index, "Balance"] = fundmgmt_debtors

                index = self.get_index("other trade debt.*")
                df2.at[index, "Balance"] -= fundmgmt_debtors

                # fundmgmt = float(input("What is the amount of trade debtors for fund management: $"))
                # index = self.get_index("trade debt.*fund m.*")[0]
                # df2.at[index,"Balance"] = fundmgmt

                # index = self.get_index("other trade debt.*")[0]
                # df2.at[index,"Balance"] -= fundmgmt
                break
            else:
                a = input(
                    "Please enter a valid input \nIs there any amount in other trade debtors that is for fund management (y/n): ").lower()
        return df2

    def rpt_a(self):
        a = input("Are there any accounts in other current assets that is amount due from director or loans to related company (y/n): ").lower()
        while a != 'n' or a != 'y':
            if a == 'n':
                break
            elif a == 'y':
                due_from_dir = input(
                    "Enter the client account numbers for amounts due from director or connected persons (separate account numbers with a comma, enter 0 if not applicable: ")
                due_from_dir = due_from_dir.replace(" ", "").split(",")
                try:
                    index_due_from_dir = self.get_index("[^\w]secure", 
                                                        start_field="due from dir", end_field="other curr.*asse")
                    df2.at[index_due_from_dir, "Balance"] = df[df["Account No"].isin(due_from_dir)].Amount.sum()

                    index_oca = self.get_index("other", start_field="other curr.*ass", end_field="total curr.*ass")

                    df2.at[index_oca,
                           "Balance"] -= df2.at[index_due_from_dir, "Balance"]
                except:
                    print(f"Input received: {due_from_dir}")
                loan = input(
                    "Enter the client account numbers for loans to related company or associated persons (separate account numbers with a comma, enter 0 if not applicable): ")
                loan = loan.replace(" ", "").split(",")
                try:
                    df2[df2["Field"].str.contains("loan.*to.*relat.*", case=False, regex=True, na=False)]
                    index_loan_to = self.get_index("loan.*to.*relat.*")
                    df2.at[index_loan_to, "Balance"] = df[df["Account No"].isin(loan)].Amount.sum()
                    df2.at[index_oca,"Balance"] -= df2.at[index_loan_to, "Balance"]
                except:
                    print(f"Input received: {loan}")
                break
            else:
                a = input("Please enter a valid input \nAre there any accounts in other current assets that is amount due from director or loans to related company (y/n): ").lower()
        return df2

    def exception_indicator(self):
        # Exception indicator if there are some L/S codes provided by YSL or TZK but no balance mapped from the client TB
        df2["L/S"].replace(0, np.nan, inplace=True)
        for i in df2.index:
            if pd.notna(df2.at[i,"L/S"]) and pd.isna(df2.at[i, "Balance"]):
                df2.at[i, "Exception indicator"] = "Not mapped"
        return df2

class ARListing:
    def __init__(self, file, sheet_name):
        self.file = file
        self.sheet_name = sheet_name
    
    def process_ar(self):
        self.df3 = pd.read_excel(self.file, sheet_name = self.sheet_name)

        for i in range(35,37):
            self.df3.at[i,"Unnamed: 3"] = self.df3.at[i, "Unnamed: 2"]

        self.df3 = self.df3.iloc[8:37, [1,3]]
        self.df3 = self.df3.dropna()
        self.df3.columns = self.df3.iloc[0]
        self.df3 = self.df3.drop(self.df3.index[0])
        return self.df3
    
    def process_ageing(self):
        self.ar_ageing = pd.read_excel(self.file, sheet_name = self.sheet_name)

        self.ar_ageing = self.ar_ageing.iloc[8:37, 1:9]
        self.ar_ageing.columns = self.ar_ageing.iloc[0]
        self.ar_ageing.drop(self.ar_ageing.index[0], inplace = True)

        pd.set_option('display.float_format', '{:.4f}'.format)

        self.ar_ageing.iloc[:23,1:] = self.ar_ageing.iloc[:23,1:].apply(self.convert)

        self.ar_ageing.drop([30,31,32,33,36], inplace = True)

        return self.ar_ageing
    
    def convert(self, x):
        pattern = "converted.*[at|@].*(\d{1}\.\d{4})"
        rate = float(re.findall(pattern, str(self.ar_ageing["Name"]))[0])
        x = pd.to_numeric(x, errors = 'coerce') * rate
        return x

if __name__ == "__main__":
    file = input("Enter the TB file path: ")
    # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\myer gold\Myer Gold Investment Management - 2022 TB.xlsx
    fy = int(input("Enter the Financial Year (e.g. 2022): "))
    tb = TBReader(file, fy)
    # df = tb.process_icm()     # only use for ICM
    df = tb.tbprocessing()

    template_fp = input("Enter mapping template file path: ")
    # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\MAS Forms mapping template - Compiled v20231027.xlsx
    f1 = TemplateReader(template_fp, sheet_name="Form 1")
    df2 = f1.templateprocessing()

    ar_fp = input("Enter AR listing file path: ")
    # file path: D:\Daciachinzq\Desktop\work\CPA FS Form 1\myer gold\Account receivables listing.xlsx
    ar = ARListing(ar_fp, sheet_name="5201 AR")

    df3 = ar.process_ar()

    g = GeneralAcc(df)

    general_accounts = df[df["L/S"] >= 6900.4].groupby("FY").agg(
        {"FY": 'first', "Account No": ",".join, "Name": ",".join, "L/S": "first", "Class": 'first', "Amount": "sum"}).reset_index(drop=True)

    general_accounts = g.one_acc(6900.2)
    general_accounts = g.one_acc(6850)
    other_ncl = [6000.3, 6000.4, 6050.2, 6100.4,
                 6200.2, 6300.2, 6350, 6400.2, 6500.2]
    general_accounts = g.aggregate(other_ncl)
    fixed_assets = [5500.1, 5500.2]
    general_accounts = g.aggregate(fixed_assets)
    other_cl = [6000, 6000.1, 6000.2, 6050, 6050.1, 6100, 6100.1, 6100.2,
                6100.3, 6150, 6200, 6200.1, 6300, 6300.1, 6400, 6400.1, 6500, 6500.1, 6600]
    general_accounts = g.aggregate(other_cl)
    general_accounts = g.one_acc(5100.1)
    general_accounts = g.one_acc(5000.0)
    inv_in_subsi = [5100.3, 5100.4, 5100.5]
    general_accounts = g.aggregate(inv_in_subsi)
    goodwill = [5700.1, 5700.1, 5800.1, 5800.2]
    general_accounts = g.aggregate(goodwill)
    other_nca = [5650.1, 5850, 5950]
    general_accounts = g.aggregate(other_nca)
    general_accounts = g.one_acc(5200.3)
    general_accounts = g.one_acc(5200.2)
    trade_debtors = [5200, 5200.1]
    general_accounts = g.aggregate(trade_debtors)

    general_accounts["L/S"] = general_accounts["L/S"].apply(
        lambda x: str("%.3f" % x))

    df2 = g.map_acc(general_accounts, df2)

    index_re = df2[df2["Field"].str.contains(
        "profit.*loss.*", case=False, regex=True, na=False)].index[0]
    df2.at[index_re, "Balance"] = general_accounts.at[0, "Amount"]

    inputs = Inputs(df, df2, df3)

    df2 = inputs.deposit()
    share_cap = inputs.pref_shares()
    share_cap = inputs.verify_share_tagging(share_cap)
    df2 = inputs.map_pref_shares(share_cap)
    df2 = inputs.trade_creditors()  # n
    df2 = inputs.rpt_l()    # y, 149782, 150512, 0
    df2 = inputs.trade_debtors()  # y, Harvest Platinium International Limited, Equity Summit Limited, Albatross Group, Nido Holdings Limited, Albatross Platinium VCC, Teo Joo Kim or Gerald Teo Tse Sian Or Teo, Oyster Enterprises Limited, Oyster Enterprises Limited, Lawrence Barki, Nico Gold Investments Ltd, UNO Capital Holdings Inc, Boulevard Worldwide Limited, Apollo Pte Limited, CAMSWARD PTE LTD, Granada Twin Investments, UNO Capital Holdings Inc, T & T Strategic Limited, Myer Gold Allocation Fund, Nasor International Limited, Tricor Services (BVI) Limited, Penny Yap, White Lotus Holdings Limited
    df2 = inputs.rpt_a()    # y, 189928, 193581, 201616 ; loans: 200155,200886
    df2 = inputs.exception_indicator()
    df2 = f1.col_map()
    df2 = f1.sum_totals()

    