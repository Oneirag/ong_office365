"""
Based on Office365-REST-Python-Client for api access and
https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
and
https://blog.darrenjrobinson.com/decoding-azure-ad-access-tokens-with-python/
Needs pip install msal msal_extensions pyjwt requests datetime
"""
import os.path
from typing import Optional

from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.files.system_object_type import FileSystemObjectType
from office365.sharepoint.listitems.listitem import ListItem
from office365.sharepoint.webs.web import Web

from ong_office365.ong_office365_base import Office365Base


class Sharepoint(Office365Base):

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
