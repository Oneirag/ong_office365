from typing import Type

from ong_office365.ong_office365_base import Office365Base
from ong_office365.ong_onedrive import OneDrive
from tests.test_ong_sharepoint import TestOngOffice365Base, iterate_client_ids


class TestOnedrive(TestOngOffice365Base):

    @staticmethod
    def client_class() -> Type[Office365Base]:
        return OneDrive

    @iterate_client_ids
    def test_101_list_files(self, client_id: str, sharepoint: OneDrive):
        """Tests that client_id can list files in the endpoint"""
        print(sharepoint.list_files())

    @iterate_client_ids
    def test_100_list_drives(self, client_id: str, sharepoint: OneDrive):
        """Tests that client_id can list files in the endpoint"""
        print(sharepoint.list_drives())
