"""
Authenticate in office365 using selenium
"""
from __future__ import annotations

import os

from ong_utils import Chrome, find_js_variable
from requests.sessions import Session
from ong_utils import decode_jwt_token

from ong_office365 import config, logger
from office365.runtime.auth.token_response import TokenResponse


def find_antiforgery_token(page_source: str) -> str | None:
    """Finds antiforgery token within page source (it is a javascript)"""
    return find_js_variable(page_source, "antiForgeryToken", ":")


class SeleniumTokenManager:

    def __init__(self):
        username = os.path.split(os.path.expanduser('~'))[-1]
        self.driver_path = None
        self.profile_path = None
        self.last_token_office = None
        self.block_pages = None

        def format_user(value, username: str):
            if not value:
                return value
            else:
                return value.format(user=username)

        if "selenium" in config.sections():
            self.driver_path = format_user(config["selenium"].get("chrome_driver_path"), username)
            # Path To Custom Profile (needed for using browser cache)
            self.profile_path = format_user(config["selenium"].get("profile_path"), username)
            self.block_pages = config["selenium"].get("block_pages")

        self.chrome = Chrome(driver_path=self.driver_path, profile_path=self.profile_path,
                             logger=logger, block_pages=self.block_pages)

    def get_auth_forms_session(self, session: Session = None) -> Session | None:
        """Returns a request.Session object with the right cookie and headers for ms forms api"""
        session = session or Session()
        cookies, antiforgery_token = SeleniumTokenManager().get_auth_forms_cookies()
        if not antiforgery_token:
            return None
        for c in cookies:
            c.pop("sameSite", None)
            c['expires'] = c.pop('expiry', None)
            c['rest'] = {'HttpOnly': c.pop('httpOnly')}
            session.cookies.set(**c)
        # Add header for validation token
        session.headers.update({'__requestverificationtoken': antiforgery_token})
        return session

    def get_auth_forms_cookies(self, timeout_headless: int =4) -> tuple:
        """
        Gets all cookies needed for ms forms api calls and gets also antiForgeryToken
        :param timeout_headless: time to wait for page load in headless mode before launching browser window
        :return: a tuple cookies_list, antiForgeryToken. Cookies_list is a list of dictionaries
        """
        """"""
        # First, attempt to get cookie from cache
        url = "https://www.office.com/login?es=Click&ru=%2F"
        # For ms forms
        url = "https://go.microsoft.com/fwlink/p/?LinkID=2115709&clcid=0x409&culture=en-us&country=us"
        url = "https://forms.office.com/landing"  # ¿Can be used for checking?
        url = "https://forms.office.com/Pages/DesignPageV2.aspx?origin=Marketing"
        cookie_name = "OIDCAuth.forms"
        driver = self.chrome.wait_for_cookie(url, cookie_name=cookie_name, timeout=60*4,
                                             timeout_headless=timeout_headless)
        if driver is None:
            raise ValueError("Could not authenticate")
        cookies_list = driver.get_cookies()
        anti_forgery = find_antiforgery_token(driver.page_source)
        self.chrome.quit_driver()
        return cookies_list, anti_forgery

    def get_auth_office(self, force_refresh: bool = False, force_logout: bool = False):
        """Gets token from https://www.office.com"""
        if self.last_token_office and not force_refresh:
            return self.last_token_office
        # Easier --- office365 main page
        logout_url = "https://www.office.com/estslogout?ru=%2F"
        url = "https://www.office.com/login?es=Click&ru=%2F"
        # url = "https://www.office.com/?auth=2"
        if force_logout:
            driver = self.chrome.get_driver(headless=True)  # Start with headless
            driver.get(logout_url)
        request_url = "sharepoint.com/_api/"
        req = self.chrome.wait_for_request(url, request_url, timeout=60, timeout_headless=10)
        if not req:
            token = None
        else:
            token = req.headers['Authorization'].split(" ")[-1]
        self.last_token_office = token
        return token

    def get_token_office(self) -> TokenResponse:
        """Authenticates to www.office.com returning token as dict that can be used with office365 library"""
        _ = self.get_auth_office()

        token_dict = dict(access_token=self.last_token_office, token_type="Bearer")
        return TokenResponse.from_json(token_dict)

    @property
    def last_decoded_token(self) -> dict:
        """Return last access token decoded as a dict"""
        if not self.last_token_office:
            self.get_auth_office(force_refresh=True)
        decoded_token = decode_jwt_token(self.last_token_office)
        return decoded_token


if __name__ == '__main__':
    token_manager = SeleniumTokenManager()
    print(token_manager.get_auth_forms_cookies())
    print(token_manager.get_auth_office())


