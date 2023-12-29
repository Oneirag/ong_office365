"""
Takes a file with current registered client_ids (downloaded from
https://go.microsoft.com/fwlink/?linkid=2083908)
and tries if those client ids can be used with sharepoint or with onedrive (microsoft graph)
"""

import sys
import os
import pandas as pd
from ong_office365.ong_sharepoint import Sharepoint
from ong_office365.ong_onedrive import OneDrive
from ong_office365 import config, logger

if sys.platform.startswith("win"):
    from pywinauto.keyboard import send_keys


class TestFunctions:
    """Wrapper for test functions"""

    @staticmethod
    def test_sharepoint(client_id):
        sp = Sharepoint(client_id, timeout=0.5)
        print(sp.me())

    @staticmethod
    def test_onedrive(client_id):
        one = OneDrive(client_id, timeout=0.5)
        print(one.list_drives())


def check_client_ids(in_file: str, out_file: str, test_func=TestFunctions.test_sharepoint):
    df = pd.read_csv(in_file)
    df_out_cols = ['display_name', "client_id", "error"]
    try:
        df_out = pd.read_csv(out_file)[df_out_cols]
        good = df_out[df_out['error'].isna()]
        for idx, row in good.iterrows():
            # print(f"'{row.client_id}': '{row.display_name[1:-1]}',")
            print(f"\t# {row.display_name}\n\t{row.client_id}")
        print(good[['display_name', 'client_id']])
    except:
        df_out = pd.DataFrame(columns=df_out_cols)

    for idx, row in df.iterrows():
        display_name = row['displayName']
        client_id = row.get("appID", row.get("appId"))
        new_data = dict(display_name=display_name, client_id=client_id)
        old_data = df_out[df_out['client_id'] == client_id]
        if old_data.empty:
            if row.get("applicationType", "") == "Microsoft Application":
                new_data['error'] = "Microsoft Application"
            else:
                try:
                    test_func(client_id)
                except Exception as e:
                    print(f"bad client: {display_name}")
                    new_data['error'] = repr(e)
                else:
                    print(f"good client: {display_name}")
                    new_data['error'] = None
                finally:
                    if sys.platform.startswith("darwin"):
                        cmd = """
                        osascript -e 'tell application "System Events" to keystroke "w" using {command down}' 
                        """
                        # minimize active window
                        os.system(cmd)
                    elif sys.platform.startswith("win"):
                        # Send Ctrl + W (for Chrome)
                        send_keys("^W")

            df_out = pd.concat([df_out, pd.DataFrame([new_data])], ignore_index=True)
            df_out.to_csv(out_file, index=False)


if __name__ == '__main__':
    in_file = "EnterpriseAppsList.csv"
    apps = ["sharepoint", "onedrive"]
    for app in apps:
        email = config("email")
        out_file = f"client_ids_{email}_{app}.csv"
        test_func = getattr(TestFunctions, f"test_{app}")
        check_client_ids(in_file, out_file, test_func=test_func)
