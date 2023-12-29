import datetime
import unittest
import os
from typing import Type

from ong_office365.ong_office365_base import Office365Base
from ong_office365.ong_sharepoint import Sharepoint
from tests.test_ong_office365_base import TestOngOffice365Base, iterate_client_ids


def get_dest_folder(sharepoint) -> str:
    """Gets remote destination folder for tests, which is the folder from the first file found"""
    folders, files = sharepoint.get_all_folders_files(limit=50)
    first_file = list(files.values())[0]
    return os.path.dirname(first_file.serverRelativeUrl)


class TestSharepoint(TestOngOffice365Base):

    single = True

    @staticmethod
    def client_class() -> Type[Office365Base]:
        return Sharepoint

    @iterate_client_ids
    def test_000_me(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can open me endpoint"""
        print(sharepoint.me())

    @iterate_client_ids
    def test_001_title(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can open site title endpoint"""
        print(sharepoint.site_title())

    @iterate_client_ids
    def test_002_personal_site(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can open personal site endpoint"""
        print(sharepoint.get_personal_site())

    @iterate_client_ids
    def test_100_list_folders(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can list folders in the endpoint"""
        folder_list = sharepoint.list_folders()
        print(folder_list)
        subfolder_list = sharepoint.list_folders(list(folder_list.keys())[-1])
        print(subfolder_list)

    @iterate_client_ids
    def test_110_list_files_in_folder(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can list files in a certain folder"""
        folder_list = sharepoint.list_folders()
        subfolder_list = sharepoint.list_folders(list(folder_list.keys())[-1])
        print(sharepoint.list_files_folder(list(subfolder_list.keys())[-1]))

    @iterate_client_ids
    def test_120_list_all_files(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can list files in the endpoint"""
        folders, files = sharepoint.get_all_folders_files()
        sharepoint.logger.info(f"Site contains: {len(files)} files and {len(folders)} folders")
        self.assertTrue(len(folders) > 0)
        self.assertTrue(len(files) > 0, "Site has no files")

    @iterate_client_ids
    def test_200_download_files(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can download files in the endpoint. Lists all files and
        downloads the largest and smallest ones"""
        all_folders, all_files = sharepoint.get_all_folders_files()
        all_files = list(all_files.values())
        sorted_files = sorted(all_files, key=lambda x: x.length)
        smallest_file = sorted_files[0]
        largest_file = sorted_files[-1]
        for file in [smallest_file, largest_file]:
            relative_url = file.serverRelativeUrl
            file_size = file.length
            file_size_mb = file_size / (1024 ** 2)
            sharepoint.logger.info(f"Downloading {relative_url} of size {file_size_mb:.2f}MB")
            filename = os.path.basename(relative_url)
            if os.path.isfile(filename):
                os.remove(filename)
            if relative_url == smallest_file:
                sharepoint.download_file(relative_url)
            else:
                sharepoint.download_file_large(relative_url)
            self.assertTrue(os.path.isfile(filename),
                            msg=f"File {relative_url} could not be downloaded")
            self.assertEqual(os.stat(filename).st_size, file.length,
                             msg=f"Wrong downloaded size for file {relative_url}")
            os.remove(filename)

    @iterate_client_ids
    def test_300_upload_files(self, client_id: str, sharepoint: Sharepoint):
        # Create a temp file and upload to sharepoint
        try:
            dest_url = self._get_configs("dest_url")[0]
        except:
            dest_url = get_dest_folder(sharepoint)
        sharepoint.logger.info(f"Uploading file to {dest_url}")
        timestamp = datetime.datetime.now().timestamp()
        temp_file = f"temporal_{timestamp}.txt"
        with open(temp_file, "w") as f:
            f.write(f"Test data created at: {timestamp}")
        try:
            sharepoint.upload_file(temp_file, dest_url)
        finally:
            os.remove(temp_file)
        file_url = dest_url + "/" + os.path.basename(temp_file)
        self.assertTrue(sharepoint.exits(file_url),
                        f"File {temp_file} was not uploaded")
        sharepoint.delete(file_url)

    @iterate_client_ids
    def test_400_list_lists(self, client_id: str, sharepoint: Sharepoint):
        """List the available sharepoint lists of current site"""
        self.verify_scopes(client_id, sharepoint, target_scopes=[r'^Sites\.(ReadWrite|FullControl)\.All$'])
        res = sharepoint.get_lists()
        print(res)
        self.assertTrue(len(res) > 2, f"Too few lists: {res}")
        pass

    @iterate_client_ids
    def test_410_read_list(self, client_id: str, sharepoint: Sharepoint):
        """List the available sharepoint lists of current site. Reads them by title, guid and object
        and checks that all return same values"""
        lists = sharepoint.get_lists()
        last_list = list(lists.values())[-1]
        data1 = sharepoint.read_list(list_obj=last_list)
        data2 = sharepoint.read_list(list_id=last_list.id)
        data3 = sharepoint.read_list(list_title=last_list.title)
        self.assertTrue(data1.equals(data2))
        self.assertTrue(data1.equals(data3))
        pass


if __name__ == '__main__':
    unittest.TestLoader.sortTestMethodsUsing = None
    unittest.main()
