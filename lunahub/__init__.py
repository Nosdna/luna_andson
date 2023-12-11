import importlib.util
import os
import luna
from luna import settings

from luna.lunahub.connection import LunaHubConnector
from luna.lunahub.connection import LunaHubBaseUploader

if True:
    # Get the config fp from luna\settings.py
    LUNAHUB_CONFIG_FILEPATH = settings.LUNAHUB_CONFIG_FILEPATH

    # Load the config via path import
    name = os.path.splitext(os.path.basename(LUNAHUB_CONFIG_FILEPATH))[0]
    spec = importlib.util.spec_from_file_location(name, LUNAHUB_CONFIG_FILEPATH)
    LUNAHUB_CONFIG = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(LUNAHUB_CONFIG)
    LUNAHUB_CONFIG = LUNAHUB_CONFIG.CONN_DICT

    # Clear vars
    del (LUNAHUB_CONFIG_FILEPATH, name, spec)

