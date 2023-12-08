'''
Contains settings for the lunaapp.

Includes:
    - LUNAHUB_CONFIG_FILEPATH -> secrets.py that contains login credentials.
    - PYEASYLIB_FOLDERPATH    -> folderpath of the pyeasylib
'''


import os
import sys

loginid = os.getlogin().lower()

###############################################################
# For developers to configure
##############################################################
if loginid == "owghimsiong":
    LUNAHUB_CONFIG_FILEPATH = r"D:\Desktop\owgs\CODES\luna\personal_workspace\db\secrets.py"
    PYEASYLIB_FOLDERPATH = r"D:\Desktop\owgs\CODES\pyeasylib"

elif loginid == "daciachinzq":    
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    PYEASYLIB_FOLDERPATH = None    #SET HERE
    
elif loginid == "gohjiawey":
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    PYEASYLIB_FOLDERPATH = None    #SET HERE


elif loginid == "phuasijia":
    LUNAHUB_CONFIG_FILEPATH = None #SET HERE
    PYEASYLIB_FOLDERPATH = None    #SET HERE

else:
    raise Exception ("Sorry. You are not authorised to run this.")
    
    
    
####################################################################
# Do not touch the codes from this point onwards.
####################################################################

# Check that LUNAHUB_CONFIG_FILEPATH and PYEASYLIB_FOLDERPATH
# are configured.
if True:
    var_to_config = []
    if LUNAHUB_CONFIG_FILEPATH is None:
        var_to_config.append('LUNAHUB_CONFIG_FILEPATH')
    if PYEASYLIB_FOLDERPATH is None:
        var_to_config.append('PYEASYLIB_FOLDERPATH')
    if len(var_to_config) > 0:
        raise Exception (
            f"Hello {loginid}!\n\nPlease set {' and '.join(var_to_config)} "
            f"at {__file__}.")
        
    del var_to_config

# Add sys.path for PYEASYLIB_FOLDERPATH
if True:
    
    # Get the folder containing pyeasylib, not pyeasylib itself.
    # In case the user key in the actual pyeasylib folderpath.
    pyeasylib_folderpath = PYEASYLIB_FOLDERPATH
    while os.path.basename(pyeasylib_folderpath) == "pyeasylib":
        
        #print ('before:',pyeasylib_folderpath )
        
        # update
        pyeasylib_folderpath = os.path.dirname(pyeasylib_folderpath)
        
        #print ('after:',pyeasylib_folderpath )
    
    # Add to sys.path
    if pyeasylib_folderpath not in sys.path:
        sys.path.append(pyeasylib_folderpath)
    
    # try to import
    import pyeasylib
    
    # Del
    del pyeasylib_folderpath
    