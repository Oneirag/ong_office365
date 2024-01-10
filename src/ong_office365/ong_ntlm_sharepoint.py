from office365.sharepoint.client_context import ClientContext, AuthenticationContext, RequestOptions
from requests_ntlm import HttpNtlmAuth
from ong_office365.ong_sharepoint import Sharepoint


class NTMLAuth(AuthenticationContext):
    def __init__(self, url: str, username: str, password: str):
        super().__init__(url)
        self.username = username
        self.password = password

    def authenticate_request(self, request):
        # type: (RequestOptions) -> None
        """Authenticate request"""
        request.auth = HttpNtlmAuth(self.username, self.password)


class NTLMSharepoint(Sharepoint):
    """
    Extends the Sharepoint class to a site that uses NTLM as authentication with username and password,
    instead of using JKT tokens
    """
    def __init__(self, base_url: str, username: str, password: str):
        """
        Creates the class, using  base_url (part of the site url before /sites), and username and password for auth
        """
        self.ctx = ClientContext(base_url, auth_context=NTMLAuth(base_url, username, password))

