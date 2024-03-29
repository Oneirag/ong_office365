from __future__ import annotations

import unittest
from typing import Type
import re

from ong_office365.ong_office365_base import Office365Base
from abc import abstractmethod
from ong_office365 import logger
from ong_office365 import test_config as config

unittest.TestLoader.sortTestMethodsUsing = None


def parse_client_id(client_id: str) -> str:
    """Parses client_id either from client id or from an url of the client, such as
    https://launcher.myapps.microsoft.com/api/signin/{client_id_goes_here}?tenantId={tenant_id_goes_here}"
    """
    if client_id.startswith("https://"):
        retval = client_id.split("?")[0].split("/")[-1]
    else:
        retval = client_id
    return retval


def iterate_client_ids(f):
    """
    Decorator for test functions iterating among all the defined client ids
    Reads self.single: True to execute just the first test, a number to execute the n-esim test,
    False (otherwise) to execute all tests
    :param f: function to decorate
    :return: decorated function
    """

    def deco(self: TestOngOffice365Base):
        values = self.clients.items()
        if self.single is not False:
            if self.single is True:
                index = 0
            else:
                index = self.single
            values = [list(values)[index]]
        for client_id, sharepoint in values:
            with self.subTest(client_id=client_id):
                f(self, client_id, sharepoint)

    return deco


class TestOngOffice365Base(unittest.TestCase):

    single: bool | int = False

    # Change in child classes
    @staticmethod
    @abstractmethod
    def client_class() -> Type[Office365Base]:
        """Gives the class to initialize for underlying tests, child of Office365Base"""
        return Office365Base

    @classmethod
    def _get_configs(cls, key) -> list:
        return config(cls.client_class().config_section()).get(key)

    @classmethod
    def client_ids(cls):
        return cls._get_configs("client_id")

    @classmethod
    def setUpClass(cls):
        client_class = cls.client_class()
        config_dict = config(cls.client_class().config_section())
        server = config_dict.get("site_url")
        email = config_dict.get("email")
        tenant = config_dict.get("tenant")
        cls.clients = {parse_client_id(client_id): client_class(client_id=parse_client_id(client_id),
                                                                email=email, server=server,
                                                                tenant=tenant,
                                                                )
                       for client_id in cls.client_ids()}

    def verify_scopes(self, client_id, client: Office365Base, target_scopes: list | str = None):
        """
        Checks that expected scopes have been received for given client
        :param client_id: id of client (just for logging purposes)
        :param client: object for getting received scopes
        :param target_scopes: list of scopes (as regular expressions) to be matched against received
        scopes
        :return: True if received scopes are valid (there are no missing ones)
        """
        scopes = client.token_scopes()
        logger.debug(f"Scopes received for {client_id}: {scopes}")
        self.assertTrue(len(scopes) > 0)
        if isinstance(target_scopes, str):
            target_scopes = [target_scopes]
        elif target_scopes is None:
            target_scopes = list()
        received_scopes = []
        for target_scope in target_scopes:
            for scope in scopes:
                if re.match(target_scope, scope):
                    received_scopes.append(target_scope)
        missing_scopes = set(target_scopes).difference(received_scopes)
        self.assertTrue(len(missing_scopes) == 0,
                        f"Some expected scopes where not received: {missing_scopes}")


if __name__ == '__main__':
    unittest.TestLoader.sortTestMethodsUsing = None
    unittest.main()
