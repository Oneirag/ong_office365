import datetime
import unittest
import os
from typing import Type

from ong_office365.ong_office365_base import Office365Base
from ong_office365.ong_sharepoint import Sharepoint
from tests.test_ong_office365_base import TestOngOffice365Base, iterate_client_ids


class TestSharepoint(TestOngOffice365Base):

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
    def test_100_list_files(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can list files in the endpoint"""
        print(sharepoint.list_files())

    @iterate_client_ids
    def test_200_download_files(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can download files in the endpoint"""
        relative_urls = [url for url in self._get_configs("relative_urls")]
        for relative_url in relative_urls:
            filename = os.path.basename(relative_url)
            if os.path.isfile(filename):
                os.remove(filename)
            # sharepoint.download_file(relative_url)
            sharepoint.download_file_large(relative_url)
            self.assertTrue(os.path.isfile(filename), f"File {relative_url} could not be downloaded")
            os.remove(filename)

    @iterate_client_ids
    def test_300_upload_files(self, client_id: str, sharepoint: Sharepoint):
        # Create a temp file and upload to sharepoint
        dest_url = self._get_configs("dest_url")[0]
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


if __name__ == '__main__':
    unittest.TestLoader.sortTestMethodsUsing = None
    unittest.main()