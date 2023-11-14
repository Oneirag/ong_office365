import sys

import msal
from msal_extensions import PersistedTokenCache, FilePersistenceWithDataProtection, KeychainPersistence, FilePersistence
from office365.runtime.auth.token_response import TokenResponse
import re

TIMEOUT_INTERACTIVE_FLOW = 20
TIMEOUT_INTERACTIVE_FLOW = 1
#TIMEOUT_INTERACTIVE_FLOW = 1
TIMEOUT_INTERACTIVE_FLOW = 0.25
#TIMEOUT_INTERACTIVE_FLOW = 20
# TIMEOUT_INTERACTIVE_FLOW = None


class MsalTokenManager:
    def __init__(self, client_id: str, email: str, server: str = None):
        self.server = server
        self.email = email
        self.tenant_prefix = self.email.split("@")[-1].replace(".", "")
        self.tenant_name = self.tenant_prefix + ".onmicrosoft.com"
        self.authority = 'https://login.microsoftonline.com/' + self.tenant_name
        self.client_id = client_id
        self.location = "token_cache.bin"
        self.scopes = self.get_scopes()
        self.persistence = self.msal_persistence()
        self.cache = PersistedTokenCache(self.persistence)

    def get_scopes(self) -> list:
        """Gets a list of scopes for auth"""
        scopes = ['.default']
        if self.server is None:  # a Graph client as default
            retval = [f"https://graph.microsoft.com/{scope}" for scope in scopes]
        else:
            pattern = r"^https://(?P<tenant>\w+(-my)?).sharepoint.com/"
            if match := re.match(pattern, self.server):
                scope_tenant = match['tenant']
                retval = [f"https://{scope_tenant}.sharepoint.com/{scope}" for scope in scopes]
            else:
                raise ValueError(f"Server {self.server} not understood")
        return retval
        # self.scopes = ["User.Read"]
        # self.scopes = ["User.Read.All"]     # Needs administrative permissions
        # self.scopes = ["User.ReadBasic.All"]     # Does not need administrative permissions
        # #self.scopes = ["User.ReadWrite"]     # Needs administrative permissions
        # #self.scopes = []     # Needs administrative permissions

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
        return result

    def msal_delegated_interactive_flow(self, scopes, prompt=None, login_hint=None, domain_hint=None, claims_challenge=None,
                                        timeout=TIMEOUT_INTERACTIVE_FLOW, port=None, extra_scopes_to_consent=None):
        print("Initiate an Interactive Flow (auth via Browser) to get AAD Access and Refresh Tokens.")

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
                print("Found account in MSAL Cache: " + account['username'])
                print("Attempting to obtain a new Access Token using the Refresh Token")
                result = self.msal_delegated_refresh(account)
                if result is None:
                    # Get a new Access Token using the Interactive Flow
                    print("Interactive Authentication required to obtain a new Access Token.")
                    result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
                else:
                    break
            # account not found in cache or there is any error getting token ... refresh token!
            if result is None or result.get("error"):
                result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
        else:
            # No accounts found in the local MSAL Cache
            # Trigger interactive authentication flow
            print("First authentication for " + self.email)
            result = self.msal_delegated_interactive_flow(self.scopes, login_hint=self.email)
        # if result.get("error"):
        #    raise ValueError(result.get("error"))
        return result

    def acquire_token_response(self) -> TokenResponse:
        result = self.acquire_token()
        return TokenResponse.from_json(result)
