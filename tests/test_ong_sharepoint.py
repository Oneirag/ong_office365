import datetime
import os.path
import unittest
from ong_office365.ong_sharepoint import Sharepoint
from ong_office365 import email, site, config
unittest.TestLoader.sortTestMethodsUsing = None


def parse_client_id(client_id: str) -> str:
    """Parses client_id either from client id or from a url of the client, such as
    https://launcher.myapps.microsoft.com/api/signin/{client_id_goes_here}?tenantId={tenant_id_goes_here}"
    """
    if client_id.startswith("https://"):
        retval = client_id.split("?")[0].split("/")[-1]
    else:
        retval = client_id
    return retval


def iterate_client_ids(f, single: bool | int = False):
    """
    Decorator for test functions iterating among all the defined client ids
    :param f: function to decorate
    :param single: True to execute just the first test, a number to execute the n-esim test,
    False (otherwise) to execute all tests
    :return: decorated function
    """
    def deco(self):
        values = self.clients.items()
        if single is not False:
            if single is True:
                index = 0
            else:
                index = single
            values = [values[index]]
        for client_id, sharepoint in values:
            with self.subTest(client_id=client_id):
                f(self, client_id, sharepoint)
    return deco


class TestOngSharepointBase(unittest.TestCase):

    email = email
    sample_site = site
    client_ids = [client for client in config.get("tests", "client_ids").splitlines() if client]

    @classmethod
    def setUpClass(cls):
        cls.clients = {parse_client_id(client_id): Sharepoint(parse_client_id(client_id),
                                                              cls.email, cls.sample_site,
                                                              )
                       for client_id in cls.client_ids}

    @iterate_client_ids
    def test_000_me(self, client_id: str, sharepoint: Sharepoint):
        """Tests that client_id can open me endpoint"""
        print(sharepoint.me())


class TestSharepointDownloads(TestOngSharepointBase):

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
        relative_urls = [url for url in config.get("tests", "relative_urls").splitlines() if url]
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
        dest_url = config.get("tests", "dest_url")
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
