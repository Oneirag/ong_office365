from __future__ import annotations

from abc import abstractmethod
from ong_office365.msal_token_manager import MsalTokenManager
from ong_office365 import config, logger as log
from tqdm import tqdm


class DownloadProgressBar(tqdm):
    """
    Adapted from https://stackoverflow.com/a/64138857
    Usage:
    with DownloadProgressBar(total=whatever_total) as t:
                object.download_media(media, filepath , progress_callback=t.update_to)
    """

    def __init__(self, total: int, incremental: bool = False, logger = None):
        self.logger = logger or log
        if total == 0:
            self.logger.warning("Total size should not be zero")
        self.incremental = incremental
        super().__init__(total=total, unit="B", unit_scale=True)

    def update_to(self, current):
        if not isinstance(current, int):
            current = len(current)
        if self.incremental:
            self.update(current)
        else:
            self.update(current - self.n)


class Office365Base:
    """
    Baseclass for office365
    """
    LARGE_FILE_SIZE = 4e6  # 4Mb

    @staticmethod
    @abstractmethod
    def config_section() -> str:
        """Returns section for parameters in config file. It must be given in child classes"""
        pass

    def __get_config(self, key: str, default_value=None):
        return config(key, default_value)

    @property
    def timeout(self):
        return self.__get_config("timeout", default_value=20)

    @property
    def tenant(self) -> str:
        return self.__get_config("tenant")

    @property
    def scopes(self) -> list:
        """List of scopes to ask for permissions. Defaults to [.default], but other list of
        permissions could be asked, such as
        ['Files.Read', 'Files.Read.All', 'Files.ReadWrite', 'Sites.Read.All']
        or ["openid", "profile", "User.Read", "Files.Read", "Files.Read.All"]"""
        scopes = ['.default']
        return scopes

    def token_scopes(self) -> list:
        """List of tokens included in the token. If asking for [.default] scope provides
        the actual tokens"""
        decoded_token = self.token_manager.last_decoded_token
        return decoded_token['scp'].split(" ")

    @property
    def email(self):
        return self.__get_config("email")

    @property
    def server(self) -> str | None:
        return self.__get_config("site_url")

    @property
    def client_id(self) -> str:
        return self.__get_config("client_id")

    def __init__(self, client_id: str = None, email: str = None, server: str | None = None,
                 tenant: str = None, init_context: callable = None,
                 to_token_response: bool = True, timeout: int = None, logger=None):
        """
        Initializes sharepoint instance
        :param client_id: List of client ids could be found in
        https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
        Defaults to config(config_key, "client_id")
        :param email: your CORPORATE email, line name@tenant. Defaults to config(config_key, "email")
        :param server: the server (e.g. for a specific teams), typically is https://{tenant}.sharepoint.com/site/{site}.
        Defaults to config(config_key, "site_url")
        :param tenant: tenant name (find it in microsoft entra ID configuration in Azure portal
        https://portal.azure.com/). Defaults to config(config_key, "tenant")
        :param init_context: a function to init context with that accepts a token
        :param to_token_response: True (default) to use acquire_token_response or false to use acquire_token
        :param timeout: time for waiting for user login. Defaults to config(config_key, "timeout")
        :param logger: an optional logger. Defaults to library default logger
        """
        self.logger = logger or log
        client_id = client_id or self.client_id
        email = email or self.email
        tenant = tenant or self.tenant
        server = server or self.server
        self.token_manager = MsalTokenManager(client_id=client_id, email=email, server=server,
                                              tenant=tenant, scopes=self.scopes, timeout=timeout or self.timeout,
                                              logger=self.logger)
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
