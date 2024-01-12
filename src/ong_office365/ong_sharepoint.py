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
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.lists.list import List
from office365.sharepoint.webs.web import Web

from ong_office365.ong_office365_base import Office365Base, DownloadProgressBar


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
                 timeout=None, logger=None):
        """
        Initializes sharepoint instance
        :param client_id: List of client ids could be found in https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
        :param email: your CORPORATE email, line name@tenant
        :param server: the server (e.g. for a specific teams), typically is https://{tenant}.sharepoint.com/site/{site}
        :param tenant: tenant name (find it in Ms Entra ID configuration)
        :param timeout: time to wait for user login
        :param logger: a logger to use instead of default library logger
        """
        super().__init__(client_id, email, server, tenant, ClientContext(server or self.server).with_access_token,
                         timeout=timeout, logger=logger)

    def __get_folder_obj(self, folder_relative_url=None) -> Folder:
        """Gets a folder object according to given relative url. Returns root folder if no url is given"""
        if folder_relative_url is None:
            folder_obj = self.ctx.web.default_document_library().root_folder
        else:
            folder_obj = self.ctx.web.get_folder_by_server_relative_url(folder_relative_url)
        return folder_obj

    def list_folders(self, folder_relative_url=None) -> dict:
        """
        Gets list of folders of a certain resource as a dict indexed by folder relative url
        :param folder_relative_url: optional parameter with the server relative URL. If None, list root folder
        :return: dict of folder objects indexed by folder server relative url
        """
        folders = self.__get_folder_obj(folder_relative_url).folders.get().execute_query()
        retval = {f.serverRelativeUrl: f for f in folders}
        return retval

    def list_files_folder(self, folder_relative_url=None):
        """
        Gets list of files of a certain folder_relative_url as a dict indexed by file relative url
        :param folder_relative_url: optional parameter with the server relative URL. If None, list root folder
        :return: dict of folder objects indexed by folder server relative url
        """
        files = self.__get_folder_obj(folder_relative_url).files.get().execute_query()
        retval = {f.serverRelativeUrl: f for f in files}
        return retval

    def get_all_folders_files(self, limit: int = None):
        """Returns a tuple of dicts of ALL folders and files of the site, indexed by relative url"""
        folders = dict()
        files = dict()
        doc_lib = self.ctx.web.default_document_library()
        items = (
            doc_lib.items.select(["FileSystemObjectType"])
            .expand(["File", "Folder"])
        )
        if limit is None:
            items = items.get_all()
        else:
            items = items.get().top(limit)
        items = items.execute_query()
        for idx, item in enumerate(items):  # type: int, ListItem
            if item.file_system_object_type == FileSystemObjectType.Folder:
                folders[item.folder.serverRelativeUrl] = item.folder
                self.logger.trace(
                    "({0} of {1})  Folder: {2}".format(
                        idx, len(items), item.folder.serverRelativeUrl
                    )
                )
            else:
                files[item.file.serverRelativeUrl] = item.file
                self.logger.trace(
                    "({0} of {1}) File: {2}".format(
                        idx, len(items), item.file.serverRelativeUrl
                    )
                )
        return folders, files

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

    def download_file_large(self, server_relative_url: str, dest_folder: str = None):
        """Downloads a file with a progress bar in the given folder (or current if None)"""
        filename = os.path.basename(server_relative_url)
        if dest_folder:
            destination = os.path.join(dest_folder, filename)
        else:
            destination = filename

        source_file = self.ctx.web.get_file_by_server_relative_path(server_relative_url)
        source_file.get().execute_query()
        with open(destination, "wb") as local_file:
            file_size = source_file.length
            with DownloadProgressBar(total=file_size, incremental=False) as t:
                source_file.download_session(local_file, t.update_to).execute_query()
        self.logger.debug("[Ok] file has been downloaded: {0}".format(destination))

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
            with DownloadProgressBar(total=os.path.getsize(local_path)) as t:
                uploaded_file = target_folder.files.create_upload_session(
                    f, size_chunk, t.update_to
                ).execute_query()

        self.logger.debug("File {0} has been uploaded successfully".format(uploaded_file.serverRelativeUrl))

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
        self.logger.debug("File has been uploaded into: {0}".format(file.serverRelativeUrl))

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

    def read_list(self, list_title: str = None, list_id: str = None, list_obj: List = None) -> pd.DataFrame:
        """
        Reads a list either with list title (it might not work if name has spaces), list id
        or the list object. Only one of the three must be informed. Returns list as a pandas DataFrame
        :param list_title: name of the list
        :param list_id: guid of the list
        :param list_obj: a list object (such one returned by get_list)
        :return:
        """

        if sum(i is not None for i in [list_title, list_id, list_obj]) != 1:
            raise ValueError("Only one parameter must be informed")

        def query_large_list(target_list):
            data = []
            # type: (List) -> None
            with DownloadProgressBar(total=target_list.item_count) as t:
                paged_items = (
                    target_list.items.paged(500, page_loaded=t.update_to).get().execute_query()
                )
            for index, item in enumerate(paged_items):  # type: int, ListItem
                data.append(item.properties)
            return data

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



