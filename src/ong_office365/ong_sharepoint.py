"""
Based on Office365-REST-Python-Client for api access and
https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
and
https://blog.darrenjrobinson.com/decoding-azure-ad-access-tokens-with-python/
Needs pip install msal msal_extensions pyjwt requests datetime
"""
import os.path
from typing import Optional

import pandas as pd
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.files.system_object_type import FileSystemObjectType
from office365.sharepoint.listitems.collection import ListItemCollection
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.lists.list import List
from office365.sharepoint.webs.web import Web

from ong_office365.ong_office365_base import Office365Base


# import urllib.parse


class Sharepoint(Office365Base):

    # Make sure I can read all lists
    # @property
    # def scopes(self) -> list:
    # Mail.Read Sites.ReadWrite.All User.Read       <- those are the right scopes for sharepoint
    #     return ['User.Read', 'User.ReadBasic.All'] #, 'Files.ReadWrite.All']
    #     return ['Sites.Read.All'] #, 'Files.ReadWrite.All']
    #     return ["Sites.FullControl.All"]
    #     return ["AllSites.FullControl"]

    @staticmethod
    def config_section() -> str:
        return "sharepoint"

    def __init__(self, client_id: str = None, email: str = None, server: str = None, tenant: str = None,
                 timeout=None):
        """
        Initializes sharepoint instance
        :param client_id: List of client ids could be found in https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
        :param email: your CORPORATE email, line name@tenant
        :param server: the server (e.g. for a specific teams), typically is https://{tenant}.sharepoint.com/site/{site}
        :param tenant: tenant name (find it in Ms Entra ID configuration)
        :param timeout: time to wait for user login
        """
        super().__init__(client_id, email, server, tenant, ClientContext(server or self.server).with_access_token,
                         timeout=timeout)

    def list_files(self):
        doc_lib = self.ctx.web.default_document_library()
        items = (
            doc_lib.items.select(["FileSystemObjectType"])
            .expand(["File", "Folder"])
            # .get_all()
            .get().top(10)
            .execute_query()
        )
        for idx, item in enumerate(items):  # type: int, ListItem
            if item.file_system_object_type == FileSystemObjectType.Folder:
                print(
                    "({0} of {1})  Folder: {2}".format(
                        idx, len(items), item.folder.serverRelativeUrl
                    )
                )
            else:
                print(
                    "({0} of {1}) File: {2}".format(
                        idx, len(items), item.file.serverRelativeUrl
                    )
                )

    def download_file(self, server_relative_url: str, path: str = None):
        """In theory...for files up to 4Mb, but 20Mb could be downloaded..."""
        filename = os.path.basename(server_relative_url)
        if path:
            destination = os.path.join(path, filename)
        else:
            destination = filename

        with open(destination, "wb") as local_file:
            file = (
                self.ctx.web.get_file_by_server_relative_url(server_relative_url)
                .download(local_file)
                .execute_query()
            )

    def download_file_large(self, server_relative_url: str, path: str = None):

        def print_download_progress(offset):
            # type: (int) -> None
            print("Downloaded '{0}' bytes...".format(offset))

        filename = os.path.basename(server_relative_url)
        if path:
            destination = os.path.join(path, filename)
        else:
            destination = filename

        source_file = self.ctx.web.get_file_by_server_relative_path(server_relative_url)
        with open(destination, "wb") as local_file:
            source_file.download_session(local_file, print_download_progress).execute_query()
        print("[Ok] file has been downloaded: {0}".format(destination))

    def get_personal_site(self):
        my_site = self.ctx.web.current_user.get_personal_site().execute_query()
        # print(my_site.url)
        return my_site.url

    def get_folder(self, target_folder=None):
        """

        :param target_folder:
        :return:
        """
        list_title = "Documents"
        if target_folder is None:
            folder = self.ctx.web.lists.get_by_title(list_title).root_folder
        else:
            folder = self.ctx.web.get_folder_by_server_relative_url(target_folder)
        return folder

    def upload_file_large(self, local_path, target_folder=None):
        """
        Uploads a local file (> 4Mb) to sharepoint in chunks
        :param local_path:
        :param target_folder: example: "Shared Documents/archive"
        :return: None
        """

        def print_upload_progress(offset):
            # type: (int) -> None
            file_size = os.path.getsize(local_path)
            print(
                "Uploaded '{0}' bytes from '{1}'...[{2}%]".format(
                    offset, file_size, round(offset / file_size * 100, 2)
                )
            )

        target_folder = self.get_folder(target_folder)
        size_chunk = 1000000  # 1Mb
        with open(local_path, "rb") as f:
            uploaded_file = target_folder.files.create_upload_session(
                f, size_chunk, print_upload_progress
            ).execute_query()

        print("File {0} has been uploaded successfully".format(uploaded_file.serverRelativeUrl))

    def upload_file(self, local_path: str, target_folder=None):
        """
        Uploads a local file to sharepoint
        :param local_path:
        :param target_folder: example: "Shared Documents/archive"
        :return: None
        """
        # If file is too big then upload chunked
        if os.path.getsize(local_path) >= self.LARGE_FILE_SIZE:
            return self.upload_file_large(local_path, target_folder)

        folder = self.get_folder(target_folder)
        with open(local_path, "rb") as f:
            file = folder.files.upload(f).execute_query()
        print("File has been uploaded into: {0}".format(file.serverRelativeUrl))

    def delete(self, file_url):
        """
        Deletes a file
        :param file_url: example = "Shared Documents/SharePoint User Guide.docx"
        :return: None
        """
        file = self.ctx.web.get_file_by_server_relative_url(file_url)
        file.delete_object().execute_query()

    def exits(self, file_url: str):
        """
        Checks if file exits
        :param file_url: example -> "Shared Documents/Financial Sample.xlsx"
        :return: True or False
        """

        def try_get_file(web, url):
            # type: (Web, str) -> Optional[File]
            try:
                return web.get_file_by_server_relative_url(url).get().execute_query()
            except ClientRequestException as e:
                if e.response.status_code == 404:
                    return None
                else:
                    raise ValueError(e.response.text)

        file = try_get_file(self.ctx.web, file_url)
        return file is not None

    def get_lists(self) -> dict:
        """Returns a dict, indexed by title, of objects representing lists of site"""
        result = (
            self.ctx.web.lists.get()
            # .select(["IsSystemList", "Title"])
            .filter("IsSystemList eq false")
            .execute_query()
        )
        return {r.title: r for r in result}

    def read_list(self, list_title: str = None, list_id: str = None, list_obj: List = None):

        if sum(i is not None for i in [list_title, list_id, list_obj]) != 1:
            raise ValueError("Only one parameter must be informed")

        def print_progress(items):
            # type: (ListItemCollection) -> None
            print("Items read: {0}".format(len(items)))

        def query_large_list(target_list):
            data = []
            # type: (List) -> None
            paged_items = (
                target_list.items.paged(500, page_loaded=print_progress).get().execute_query()
            )
            for index, item in enumerate(paged_items):  # type: int, ListItem
                data.append(item.properties)
                # print("{0}: {1}".format(index, item.id))
            # all_items = [item for item in paged_items]
            # print("Total items count: {0}".format(len(all_items)))
            return data

        def get_total_count(target_list):
            # type: (List) -> None
            all_items = target_list.items.get_all(5000, print_progress).execute_query()
            print("Total items count: {0}".format(len(all_items)))

        if list_obj is not None:
            large_list = list_obj
        elif list_id is not None:
            large_list = self.ctx.web.lists.get_by_id(list_id)
        elif " " not in list_title:
            large_list = self.ctx.web.lists.get_by_title(list_title)
        else:
            # list title with spaces. Must be manually read from the list of all tables
            lists = self.get_lists()
            for name, list_obj in lists.items():
                if name == list_title:
                    return self.read_list(list_obj=list_obj)
            raise ValueError(f"List {list_title} not found")
        retval = query_large_list(large_list)
        df = pd.DataFrame(retval)
        df = df.set_index("ID")
        return df
