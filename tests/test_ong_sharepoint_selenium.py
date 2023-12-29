import unittest

from tests.test_ong_sharepoint import TestSharepoint, Office365Base, Type
from ong_office365.ong_selenium_sharepoint import SeleniumSharepoint


class TestSharepointSelenium(TestSharepoint):
    single = True

    @staticmethod
    def client_class() -> Type[Office365Base]:
        return SeleniumSharepoint

    @classmethod
    def client_ids(cls):
        return ['test_client_id']


if __name__ == '__main__':
    unittest.main()

