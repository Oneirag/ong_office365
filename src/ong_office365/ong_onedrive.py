"""
Adaptation of samples of office365.onedrive to use msal token cache
Uses ms graph. Try what can be done with ms graph in
https://developer.microsoft.com/en-us/graph/graph-explorer
"""
from ong_office365.ong_office365_base import Office365Base
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.graph_client import GraphClient


class OneDrive(Office365Base):

    @staticmethod
    def config_section() -> str:
        return "onedrive"

    def __init__(self, client_id: str = None, email: str = None, tenant: str = None, server=None,
                 timeout=None):
        server = None  # server is not needed in Graph clients, such as Onedrive
        super().__init__(client_id=client_id, email=email, server=server, tenant=tenant,
                         init_context=GraphClient, to_token_response=False, timeout=timeout)

    def drives(self):
        drives = self.ctx.drives.get().top(100).execute_query()
        for drive in drives:
            print("Drive url: {0}".format(drive.web_url))

    def me(self):
        me = self.ctx.me.get().execute_query()
        print(me.user_principal_name)
        return me.user_principal_name

    def list_files(self, max=5):

        def enum_folders_and_files(root_folder):
            # type: (DriveItem) ->  None
            drive_items = root_folder.children.get().top(max).execute_query()
            for drive_item in drive_items:
                print("Name: {0}".format(drive_item.web_url))
                if drive_item.is_folder:  # is folder facet?
                    enum_folders_and_files(drive_item)

        root = self.ctx.me.drive.root
        enum_folders_and_files(root)

    def list_drives(self):
        drives = self.ctx.drives.get().top(100).execute_query()
        for drive in drives:
            print("Drive url: {0}".format(drive.web_url))
