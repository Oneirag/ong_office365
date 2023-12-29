"""
Authenticate in office365 using selenium
"""
from __future__ import annotations

import datetime
import os

from ong_utils import Chrome, find_js_variable, InternalStorage
from requests.sessions import Session
from ong_utils import decode_jwt_token, decode_jwt_token_expiry

from ong_office365 import config, logger
from office365.runtime.auth.token_response import TokenResponse


def is_cookie_expired(cookies_list: list | None, cookie_name="OIDCAuth.forms") -> bool:
    """Checks if a certain cookie is expired. Returns True if cookie is expired or cookies_list is None"""
    if cookies_list:
        for ck in cookies_list:
            if ck['name'] == cookie_name:
                return ck['expiry'] < (datetime.datetime.now().timestamp() + 60)
    return True


def find_antiforgery_token(page_source: str) -> str | None:
    """Finds antiforgery token within page source (it is a javascript)"""
    return find_js_variable(page_source, "antiForgeryToken", ":")


def save_form_auth(cookies: list, antiforgery_token: str, session: Session):
    """Saves cookies and antiforgery token in session"""
    for c in cookies:
        c.pop("sameSite", None)
        c['expires'] = c.pop('expiry', None)
        c['rest'] = {'HttpOnly': c.pop('httpOnly', None)}
        session.cookies.set(**c)
    # Add header for validation token
    session.headers.update({'__requestverificationtoken': antiforgery_token})


class SeleniumTokenManager:

    key_auth_forms = "auth_forms"
    key_jwt_token = "jwt_token"
    # Cookie used for authentication in ms forms
    cookie_name = "OIDCAuth.forms"

    def __init__(self):
        username = os.path.split(os.path.expanduser('~'))[-1]
        self.internal_storage = InternalStorage(self.__class__.__name__)
        self.driver_path = None
        self.profile_path = None
        self.last_token_office = None
        self.block_pages = None

        def format_user(value, username: str):
            if not value:
                return value
            else:
                return value.format(user=username)

        self.driver_path = format_user(config("selenium").get("chrome_driver_path"), username)
        # Path To Custom Profile (needed for using browser cache)
        self.profile_path = format_user(config("selenium").get("profile_path"), username)
        self.block_pages = config("selenium").get("block_pages")

        self.chrome = Chrome(driver_path=self.driver_path, profile_path=self.profile_path,
                             logger=logger, block_pages=self.block_pages)

    def get_auth_forms_session(self, session: Session = None, timeout_headless=4) -> Session | None:
        """
        Returns a request.Session object with the right cookie and headers for ms forms api
        :param session: a requests.Session object to reuse. If None, a new one will be created
        :param timeout_headless: time to wait for page load in headless mode before launching browser window
        :return: requests.Session object with the right cookie and headers for ms forms api or None if could not
        authenticate
        """
        session = session or Session()
        cookies, antiforgery_token = self.__get_auth_forms_cookies()
        if not antiforgery_token:
            return None
        save_form_auth(cookies, antiforgery_token, session)
        return session

    def __check_cache_forms(self) -> tuple:
        """Checks that forms cache is valid, ensuring cookies are not expired
        and double-checks it navigating to organizationInfo"""
        cookies_list, anti_forgery = self.internal_storage.get_value(self.key_auth_forms) or (None, None)
        if not is_cookie_expired(cookies_list, self.cookie_name):
            # session = Session()
            # save_form_auth(cookies_list, anti_forgery, session)
            # req = session.get("https://forms.office.com/formapi/api/organizationInfo")
            # req = session.get("https://forms.office.com/formapi/api/forms")
            #if req.status_code != 403:      # Forbidden
            #    logger.info("Using cached forms auth")
            #    return cookies_list, anti_forgery
            logger.info("Using cached forms auth")
            return cookies_list, anti_forgery
        return None, None

    def __get_auth_forms_cookies(self, timeout_headless: int = 4) -> tuple:
        # First, attempt to get cookie and anti-forgery token from cache
        cookies_list, anti_forgery = self.__check_cache_forms()
        if not cookies_list:
            # url = "https://www.office.com/login?es=Click&ru=%2F"
            # For ms forms
            url = "https://forms.office.com/Pages/DesignPageV2.aspx?origin=Marketing"
            driver = self.chrome.wait_for_cookie(url, cookie_name=self.cookie_name, timeout=60*4,
                                                 timeout_headless=timeout_headless)
            if driver is None:
                raise ValueError("Could not authenticate")
            cookies_list = driver.get_cookies()
            anti_forgery = find_antiforgery_token(driver.page_source)
            self.chrome.quit_driver()
            self.internal_storage.store_value(self.key_auth_forms, (cookies_list, anti_forgery))
        return cookies_list, anti_forgery

    def get_auth_office(self, force_refresh: bool = False, force_logout: bool = False) -> str | None:
        """Gets token from https://www.office.com"""
        if self.last_token_office and not force_refresh:
            return self.last_token_office
        name_key = "upn"    # key to use in decoded jwt token for username
        token = self.internal_storage.get_value(self.key_jwt_token)
        if token:
            if decode_jwt_token_expiry(token) > datetime.datetime.now():
                session = Session()
                decoded = decode_jwt_token(token)
                logger.info(f"Using cached token for user '{decoded[name_key]}'")
                return token
        # Easier --- office365 main page
        logout_ufrl = "https://www.office.com/estslogout?ru=%2F"
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
            self.internal_storage.store_value(self.key_jwt_token, token)
            decoded = decode_jwt_token(token)
            logger.info(f"New token obtained for user '{decoded[name_key]}'")
        self.last_token_office = token
        self.chrome.close_driver()
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

    def clear_cache(self):
        self.internal_storage.remove_stored_value(self.key_jwt_token)
        self.internal_storage.remove_stored_value(self.key_auth_forms)


if __name__ == '__main__':
    token_manager = SeleniumTokenManager()
    token_manager.clear_cache()
    print(token_manager.get_auth_forms_session().cookies)
    print(token_manager.get_auth_office())


