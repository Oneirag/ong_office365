"""
Adaptation of samples of office365.onedrive to use msal token cache
"""
from ong_office365.ong_office365_base import Office365Base
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.graph_client import GraphClient


class OneDrive(Office365Base):
    def __init__(self, client_id, email, server):
        super().__init__(client_id=client_id, email=email, server=server, init_context=GraphClient,
                         to_token_response=False)

    def drives(self):
        drives = self.ctx.drives.get().top(100).execute_query()
        for drive in drives:
            print("Drive url: {0}".format(drive.web_url))

    def me(self):
        return self.ctx.me.execute_query()

    def list_files(self):

        def enum_folders_and_files(root_folder):
            # type: (DriveItem) ->  None
            drive_items = root_folder.children.get().execute_query()
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
