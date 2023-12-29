import sys
import os
import configparser
import loguru

from ong_utils import OngConfig

name = "ong_office365"
_cfg = OngConfig(name, default_app_cfg={
        "email": "someone@contoso.com",
        "tenant": "contoso",
        # client_id should come from https://go.microsoft.com/fwlink/?linkid=2083908
        # This is a sample value from a google search
        "client_id": "6731de76-14a6-49ae-97bc-6eba6914391e",
        "sharepoint": "https://contoso.sharepoint.com/sites/example_site",
    })
config = _cfg.config
test_config = _cfg.config_test
# logger = _cfg.logger
logger = loguru.logger


def get_email(default: str = None) -> str:
    if default:
        return default

    if sys.platform.startswith("win"):
        import win32com.client as win32
        outlook = win32.gencache.EnsureDispatch("outlook.application")
        return outlook.Sessiol.Accounts.Item(1).DisplayName
    else:
        return input("Email address: ")
