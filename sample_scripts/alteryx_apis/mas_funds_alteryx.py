# This file is to be run through cmd line.
# Input parameters to be specified through args
# Output will be via a file.

# Import standard package
import os
import sys
import argparse
import importlib.util
import time

# Set luna path - Load from settings.py
if True:
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

# Import luna packages
import luna
import luna.common as common
import luna.fsvi as fsvi
import luna.lunahub as lunahub
from luna.fsvi.funds.invmt_report_formatter import InvmtOutputFormatter

# Import help lib
import pyeasylib


# To run on command prompt
if __name__ == "__main__":
    
    # Specify the cmd line arguments requirement    
    parser = argparse.ArgumentParser()
    parser.add_argument("--client_number", required=True)
    parser.add_argument("--client_fy", required=True)
    parser.add_argument("--aic_name", required=True)
    
    # Parse the information
    if False:
        args = parser.parse_args()    
        client_number = args.client_number
        fy = args.client_fy
        aic_name = args.aic_name
        
    #############################################
    ## FOR DEBUGGING ONLY ##
    if True:
        client_number   = 70000
        fy              = 2023
        aic_name        = "yipjunhorng"
    #############################################

    # Get the luna folderpath 
    luna_init_file = luna.__file__
    luna_folderpath = os.path.dirname(luna_init_file)

    portfolio_mapper_fp = r"D:\workspace\luna\parameters\invmt_portfolio_mapper.xlsx"

    client_class = lunahub.tables.client.ClientInfoLoader_From_LunaHub(client_number)
    sublead_class = lunahub.tables.fs_funds_invmt_output_sublead.FundsSublead_DownloaderFromLunaHub(client_number, fy)
    portfolio_class = lunahub.tables.fs_funds_invmt_output_portfolio.FundsPortfolio_DownloaderFromLunaHub(client_number, fy)
    recon_class = lunahub.tables.fs_funds_invmt_txn_recon_details.FundsInvmtTxnReconDetail_DownloaderFromLunaHub(client_number, fy)
    broker_class    = lunahub.tables.fs_funds_broker_statement.FundsBrokerStatement_DownloaderFromLunaHub(client_number, fy)
    custodian_class = lunahub.tables.fs_funds_custodian_confirmation.FundsCustodianConfirmation_DownloaderFromLunaHub(client_number, fy)
    tb_class = common.TBLoader_From_LunaHub(client_number, fy)

    for attempt in range(12):
        user_response_class = lunahub.tables.fs_funds_userresponse.FundsUserResponse_DownloaderFromLunaHub(
            client_number,
            fy)
        user_inputs = user_response_class.main()
        if user_inputs is not None:
            break
        elif user_inputs is None and attempt == 11:
            raise Exception (f"Data not found for specified client {client_number} or FY {fy}.")
        else:
            continue

    output_fn = f"mas_funds_investment_{client_number}_{fy}.xlsx"
    output_fp = os.path.join(settings.TEMP_FOLDERPATH, output_fn)
    #output_fp = pyeasylib.check_filepath(output_fp)
    pyeasylib.create_folder_for_filepath(output_fp)    

    self = InvmtOutputFormatter(sublead_class   = sublead_class,
                                 portfolio_class= portfolio_class,
                                 recon_class    = recon_class,
                                 broker_class   = broker_class,
                                 custodian_class= custodian_class,
                                 tb_class       = tb_class,
                                 output_fp      = output_fp,
                                 mapper_fp      = portfolio_mapper_fp,
                                 user_inputs    = user_inputs,
                                 client_class   = client_class,
                                 fy             = fy,
                                 aic_name       = aic_name
                                 )


    print (f"Final output saved to {output_fp}.")

    # Open output file
    if True:
        import webbrowser
        webbrowser.open(output_fp)
    