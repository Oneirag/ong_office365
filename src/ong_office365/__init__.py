import sys
import os
import configparser
import loguru

logger = loguru.logger

name = "ong_office365"


def get_email(default: str = None) -> str:
    if default:
        return default

    if sys.platform.startswith("win"):
        import win32com.client as win32
        outlook = win32.gencache.EnsureDispatch("outlook.application")
        return outlook.Sessiol.Accounts.Item(1).DisplayName
    else:
        return input("Email address: ")


def init_config() -> configparser.ConfigParser:

    default = {
        "email": "someone@contoso.com",
        "tenant": "contoso",
        # client_id should come from https://go.microsoft.com/fwlink/?linkid=2083908
        # This is a sample value from a google search
        "client_id": "6731de76-14a6-49ae-97bc-6eba6914391e",
        "sharepoint": "https://contoso.sharepoint.com/sites/example_site",
    #    "onedrive": None,   # no need for a value here
    }


    config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation(),
                                       )
    config['DEFAULT'] = default

    # overwrite default values with values in home, in this folder or in current folder
    config.read([os.path.expanduser(f'~/.{name}'),
                 os.path.join(os.path.dirname(__file__), f'{name}.cfg'),
                 f'{name}.cfg',
                 ],
                encoding='utf-8')

    return config


config = init_config()