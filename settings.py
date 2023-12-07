# Load LunaHub configs
# secrets.py file is stored separately
import os

loginid = os.getlogin().lower()
if loginid == "owghimsiong":
    LUNAHUB_CONFIG_FILEPATH = r"D:\Desktop\owgs\CODES\luna\personal_workspace\db\secrets.py"


elif loginid == "daciachinzq":    
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    raise Exception (f"Please set LUNAHUB_CONFIG_FILEPATH at {__file__} for user={loginid}.")

    
elif loginid == "gohjiawey":
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    raise Exception (f"Please set LUNAHUB_CONFIG_FILEPATH at {__file__} for user={loginid}.")


elif loginid == "phuasijia":
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    raise Exception (f"Please set LUNAHUB_CONFIG_FILEPATH at {__file__} for user={loginid}.")

    
else:
    raise Exception ("Sorry. You are not authorised to run this.")