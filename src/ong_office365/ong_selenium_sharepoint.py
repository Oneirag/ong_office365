from ong_office365 import logger
from ong_office365.ong_sharepoint import Sharepoint
from ong_office365.selenium_token.office365_selenium import SeleniumTokenManager
from office365.sharepoint.client_context import ClientContext


class SeleniumSharepoint(Sharepoint):
    """Same as Sharepoint, but gets token using regular browser and selenium instead of msal with client_id"""

    def __get_decoded(self, key: str):
        """Gets a certain value from last decoded token"""
        decoded = self.token_manager.last_decoded_token
        return decoded[key]

    @property
    def email(self):
        return self.__get_decoded('upn')

    @property
    def server(self) -> str | None:
        try:
            server = super().server
            return server
        except:
            return self.__get_decoded("aud")

    def __init__(self, client_id: str = None, server: str = None):
        """Inits class. Ignores parent init and client_id (used for compatibility. Server is an optional
        parameter that can be also used"""
        self.token_manager = SeleniumTokenManager()
        self.logger = logger
        self.ctx = ClientContext(server or self.server).with_access_token(self.token_manager.get_token_office)


if __name__ == '__main__':
    ss = SeleniumSharepoint()
    print(ss.me())
