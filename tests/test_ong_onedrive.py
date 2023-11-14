from ong_office365 import config
from ong_office365.ong_onedrive import OneDrive
from tests.test_ong_sharepoint import TestOngSharepointBase, parse_client_id, iterate_client_ids


class TestOnedrive(TestOngSharepointBase):
    @classmethod
    def setUpClass(cls):
        cls.sample_site = config['DEFAULT'].get("onedrive") or None
        cls.clients = {parse_client_id(client_id): OneDrive(parse_client_id(client_id),
                                                              cls.email, cls.sample_site,
                                                              )
                       for client_id in cls.client_ids}

    @iterate_client_ids
    def test_101_list_files(self, client_id: str, sharepoint: OneDrive):
        """Tests that client_id can list files in the endpoint"""
        print(sharepoint.list_files())

    @iterate_client_ids
    def test_100_list_drives(self, client_id: str, sharepoint: OneDrive):
        """Tests that client_id can list files in the endpoint"""
        print(sharepoint.list_drives())
