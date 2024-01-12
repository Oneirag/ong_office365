"""
Based on:
https://blog.darrenjrobinson.com/interactive-authentication-to-microsoft-graph-using-msal-with-python-and-delegated-permissions/
Needs pip install msal msal_extensions pyjwt==1.7.1 requests datetime
Adapted for pyjwt 2.x using https://blog.darrenjrobinson.com/decoding-azure-ad-access-tokens-with-python/
"""
from __future__ import annotations

import re
import sys

import msal
from msal_extensions import PersistedTokenCache, FilePersistenceWithDataProtection, KeychainPersistence, FilePersistence
from office365.runtime.auth.token_response import TokenResponse
from ong_utils import decode_jwt_token

from ong_office365 import logger as log


def is_uuid(tenant) -> bool:
    """True it a tenant is a  uuid (and not .microsoftonline.com should be added for the authority"""
    UUID_PATTERN = re.compile(r'^[\da-f]{8}-([\da-f]{4}-){3}[\da-f]{12}$', re.IGNORECASE)
    return bool(UUID_PATTERN.match(tenant))


class MsalTokenManager:
    def __init__(self, client_id: str, email: str, server: str | None, tenant: str,
                 scopes: list = None, timeout: int = None, logger=None):
        """
        Initializes a token manager
        :param client_id: app client id
        :param email: user email for delegated authentication
        :param server: server to access (ms graph url, sharepoint url)
        :param tenant: tenant name or id
        :param scopes: scopes to request access to, defaults to ['.default']
        :param timeout: timeout to wait for interactive flow. Defaults to 20 sec
        :param logger: optional logger, or use library default logger
        """
        self.logger = logger or log
        self.__last_scopes = None  # scopes received in token (only for fresh tokens)
        self.__last_token = None  # Last obtained token
        self.server = server
        self.email = email
        self.tenant_prefix = tenant
        self.timeout = timeout or 20
        if is_uuid(self.tenant_prefix):
            self.tenant_name = self.tenant_prefix
        # Use common for personal Microsoft accounts and work/school accounts from Azure Active Directory
        # Use organizations for work/school accounts from Azure Active Directory only
        # Use consumers for personal Microsoft accounts (MSA) only
        elif self.tenant_prefix in ["common", "organizations", "consumers"]:
            self.tenant_name = self.tenant_prefix
        else:
            self.tenant_name = self.tenant_prefix + ".onmicrosoft.com"
        self.authority = 'https://login.microsoftonline.com/' + self.tenant_name
        self.client_id = client_id
        self.location = "token_cache.bin"
        self.scopes = self.get_scopes(scopes or ['.default'])
        self.persistence = self.msal_persistence()
        self.cache = PersistedTokenCache(self.persistence)

    @property
    def last_token(self) -> str:
        """Return last access token"""
        if not self.__last_token:
            self.acquire_token()
        return self.__last_token

    @property
    def last_decoded_token(self) -> dict:
        """Return last access token decoded as a dict"""
        decoded_token = decode_jwt_token(self.last_token)
        return decoded_token

    def get_scopes(self, scopes: list = None) -> list:
        """Gets a list of scopes for auth"""
        scopes = scopes or ['.default']
        if self.server is None:  # a Graph client as default
            retval = [f"https://graph.microsoft.com/{scope}" for scope in scopes]
        else:
            pattern = r"^https://(?P<tenant>\w+(-my)?).sharepoint.com/"
            if match := re.match(pattern, self.server):
                scope_tenant = match['tenant']
                retval = [f"https://{scope_tenant}.sharepoint.com/{scope}" for scope in scopes]
            else:
                retval = [f"{self.server}/{scope}" for scope in scopes]
                # raise ValueError(f"Server {self.server} not understood")
        self.logger.debug(f"{scopes=}")
        return retval

    def msal_cache_accounts(self, username=None):
        app = msal.PublicClientApplication(client_id=self.client_id, authority=self.authority, token_cache=self.cache)
        accounts = app.get_accounts(username)
        return accounts

    def msal_persistence(self):
        """Build a suitable persistence instance based your current OS"""
        if sys.platform.startswith('win'):
            return FilePersistenceWithDataProtection(self.location)
        if sys.platform.startswith('darwin'):
            return KeychainPersistence(self.location, "my_service_name", "my_account_name")
        return FilePersistence(self.location)

    def msal_delegated_refresh(self, account):
        app = msal.PublicClientApplication(
            client_id=self.client_id, authority=self.authority, token_cache=self.cache)
        result = app.acquire_token_silent_with_error(
            scopes=self.scopes, account=account)
        if result is not None and "error" in result:
            self.logger.debug(f"Error in token: error='{result.get('error')}' suberror='{result.get('suberror')}'")
        return result

    def msal_delegated_interactive_flow(self, scopes, prompt=None, login_hint=None, domain_hint=None,
                                        claims_challenge=None,
                                        timeout=None, port=None, extra_scopes_to_consent=None):
        self.logger.debug("Initiate an Interactive Flow (auth via Browser) to get AAD Access and Refresh Tokens.")
        timeout = timeout or self.timeout
        app = msal.PublicClientApplication(client_id=self.client_id, authority=self.authority, token_cache=self.cache)

        success_template = """<html><body><script>setTimeout(function(){window.close()}, 3000);</script></body></html>"""
        welcome_template = """<html><body><script>setTimeout(function(){window.close()}, 10000);</script></body></html>"""
        welcome_template = None
        # success_template = None
        result = app.acquire_token_interactive(scopes=scopes, login_hint=login_hint, prompt=prompt,
                                               domain_hint=domain_hint, claims_challenge=claims_challenge,
                                               timeout=timeout, port=port, success_template=success_template,
                                               error_template=success_template,
                                               welcome_template=welcome_template,
                                               extra_scopes_to_consent=extra_scopes_to_consent)
        return result

    def acquire_token(self) -> dict:
        accounts = self.msal_cache_accounts(self.email)
        result = None
        if accounts:
            for account in accounts:
                self.logger.debug("Found account in MSAL Cache: " + account['username'])
                self.logger.debug("Attempting to obtain a new Access Token using the Refresh Token")
                result = self.msal_delegated_refresh(account)
                if result is None:
                    # Get a new Access Token using the Interactive Flow
                    self.logger.debug("Interactive Authentication required to obtain a new Access Token.")
                    result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
                else:
                    break
            # account not found in cache or there is any error getting token ... refresh token!
            if result is None or result.get("error"):
                result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
        else:
            # No accounts found in the local MSAL Cache
            # Trigger interactive authentication flow
            self.logger.debug("First authentication for " + self.email)
            result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
        if result:
            if "error" in result:
                self.logger.error(f"{result['error']}: {result['error_description']}")
            else:
                self.__last_token = result['access_token']
                # Scopes are only received for fresh tokens. If token came from cache this is not received
                self.__last_scopes = result.get('scopes')
        return result

    def acquire_token_response(self) -> TokenResponse:
        result = self.acquire_token()
        return TokenResponse.from_json(result)
