import unittest
from typing import Type

from ong_office365.ong_office365_base import Office365Base
from abc import abstractmethod
from ong_office365 import config

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


class TestOngOffice365Base(unittest.TestCase):

    # Change in child classes
    @staticmethod
    @abstractmethod
    def client_class() -> Type[Office365Base]:
        """Gives the class to initialize for underlying tests, child of Office365Base"""
        return Office365Base

    @classmethod
    def _get_configs(cls, key) -> list:
        return config.get(cls.client_class().config_section(), key).strip().splitlines()

    @classmethod
    def client_ids(cls):
        return cls._get_configs("client_id")

    @classmethod
    def setUpClass(cls):
        client_class = cls.client_class()
        cls.clients = {parse_client_id(client_id): client_class(parse_client_id(client_id))
                       for client_id in cls.client_ids()}


if __name__ == '__main__':
    unittest.TestLoader.sortTestMethodsUsing = None
    unittest.main()
