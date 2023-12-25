"""
Authenticate in office365 using selenium
"""
from __future__ import annotations

import os
import re
from seleniumwire import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from requests.sessions import Session
from ong_utils import decode_jwt_token

from ong_office365 import config, logger
from office365.runtime.auth.token_response import TokenResponse


def find_antiforgery_token(page_source: str) -> str | None:
    """Finds antiforgery token within page source (it is a javascript)"""
    pattern = fr'"antiForgeryToken":"?(?P<value>.*?)"?,'
    found = re.findall(pattern, page_source)
    if found:
        return found[0]
    return None


class SeleniumTokenManager:

    def __init__(self):
        username = os.path.split(os.path.expanduser('~'))[-1]
        self.driver_path = None
        self.profile_path = None
        self.last_token_office = None

        def format_user(value, username: str):
            if not value:
                return value
            else:
                return value.format(user=username)

        if "selenium" in config.sections():
            self.driver_path = format_user(config["selenium"].get("chrome_driver_path"), username)
            # Path To Custom Profile (needed for using browser cache)
            self.profile_path = format_user(config["selenium"].get("profile_path"), username)

    def get_driver(self, headless: bool = False, avoid_home_page: bool=True):
        """Initializes a driver, either headless (hidden) or not"""
        # options = uc.ChromeOptions()
        options = webdriver.ChromeOptions()
        # Avoid annoying messages on chrome startup
        options.add_argument("--disable-notifications")
        options.add_argument("google-base-url=about:blank")
        if self.driver_path:
            options.binary_location = self.driver_path
        if self.profile_path:
            options.add_argument(f"user-data-dir={self.profile_path}")
        if headless:
            options.add_argument("--headless=new")  # for Chrome >= 109
        logger.debug(f"Initializing driver with options: {options}")
        driver = webdriver.Chrome(options=options)
        # Stop loading home page!
        if avoid_home_page:
            driver.get("chrome://version/")
            driver.execute_script("window.stop();")
        return driver

    def get_auth_forms(self) -> str | None:
        """Gets cookie value for authenticating in forms"""
        cookies = self.get_auth_forms_cookies()
        for ck in cookies:
            if ck['name'] == "OIDCAuth.forms":
                return ck['value']
        return None

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

        driver = self.get_driver(headless=True)
        driver.get(url)
        driver.implicitly_wait(time_to_wait=timeout_headless)
        ck = driver.get_cookie("OIDCAuth.forms")
        if ck:
            logger.debug("Valid cookie found in cache")
        else:
            logger.debug("No valid cooke found. Attempting manually")
            driver.quit()   # Close headless driver
            driver = self.get_driver(headless=False)
            driver.get(url)
            cookie = None
            try:
                cookie = WebDriverWait(driver, timeout=180).until(lambda d: d.get_cookie('OIDCAuth.forms'))
            except:
                logger.error("Could not find auth cookie. Ms forms authentication failed")
        cookies_list = driver.get_cookies()
        anti_forgery = find_antiforgery_token(driver.page_source)
        driver.quit()
        return cookies_list, anti_forgery

    def get_auth_office(self, force_refresh: bool = False, force_logout: bool = False):
        """Gets token from https://www.office.com"""
        if self.last_token_office and not force_refresh:
            return self.last_token_office
        driver = self.get_driver(headless=True)     # Start with headless
        # Easier --- office365 main page
        logout_url = "https://www.office.com/estslogout?ru=%2F"
        url = "https://www.office.com/login?es=Click&ru=%2F"
        # url = "https://www.office.com/?auth=2"
        # Capture only calls to sharepoint
        driver.scopes = [
            '.*sharepoint.*',
        ]
        if force_logout:
            driver.get(logout_url)
        driver.get(url)
        token = None
        try:
            req = driver.wait_for_request("sharepoint.com/_api/", timeout=4)
            token = req.headers['Authorization'].split(" ")[-1]
        except TimeoutException:
                # Retry with interactive
                driver.quit()
                driver = self.get_driver(headless=False)
                driver.get(url)
                req = driver.wait_for_request("sharepoint.com/_api/", timeout=60)
                token = req.headers['Authorization'].split(" ")[-1]
        finally:
            driver.quit()
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
    # print(token_manager.get_auth_forms())
    print(token_manager.get_auth_office())


