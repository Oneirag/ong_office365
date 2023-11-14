import sys
import os
import configparser

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


config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation(),
                                   )
config.read([os.path.expanduser(f'~/.{name}'),
             os.path.join(os.path.dirname(__file__), f'{name}.cfg'),
             f'{name}.cfg',
             ],
            encoding='utf-8')
site = config.get("DEFAULT", "sharepoint")
email = get_email(config.get("DEFAULT", "email"))
