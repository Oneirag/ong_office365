
from ong_office365.msal_token_manager import MsalTokenManager
# from ong_sharepoint.selenium_token_manager import MsalTokenManager


class Office365Base:
    """
    Baseclass for office365
    """
    LARGE_FILE_SIZE = 4e6  # 4Mb

    def __init__(self, client_id: str, email: str, server: str, init_context, to_token_response: bool=True):
        """
        Initializes sharepoint instance
        :param client_id: List of client ids could be found in https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
        :param email: your CORPORATE email, line name@tenant
        :param server: the server (e.g. for a specific teams), typically is https://{tenant}.sharepoint.com/site/{site}
        :param server: the server (e.g. for a specific teams), typically is https://{tenant}.sharepoint.com/site/{site}
        :param init_context: a function to init context with that accepts a token
        :param to_token_response: True (default) to use acquire_token_response or false to use acquire_token
        """
        self.client_id = client_id
        self.email = email
        self.server = server
        self.token_manager = MsalTokenManager(self.client_id, self.email, server=self.server)
        if to_token_response:
            self.ctx = init_context(self.token_manager.acquire_token_response)
        else:
            self.ctx = init_context(self.token_manager.acquire_token)

    def me(self):
        me = self.ctx.web.current_user.get().execute_query()
        return me.login_name

    def site_title(self) -> str:
        web = self.ctx.web.get().execute_query()
        return web.title