"""
Authenticate in office365 using selenium
"""
import os
from seleniumwire import webdriver
from selenium.webdriver.support.ui import WebDriverWait

import json
import gzip
from ong_office365 import config, logger


class SeleniumTokenManager:

    def __init__(self):
        username = os.path.split(os.path.expanduser('~'))[-1]
        self.driver_path = None
        self.profile_path = None

        def format_user(value, username: str):
            if not value:
                return value
            else:
                return value.format(user=username)

        if "selenium" in config.sections():
            self.driver_path = format_user(config["selenium"].get("chrome_driver_path"), username)
            # Path To Custom Profile (needed for using browser cache)
            self.profile_path = format_user(config["selenium"].get("profile_path"), username)

    def get_driver(self, headless: bool = False):
        """Initializes a driver, either headless (hidden) or not"""
        # options = uc.ChromeOptions()
        options = webdriver.ChromeOptions()
        # Avoid anoying messages on chrome startup
        options.add_argument("--disable-notifications")
        if self.driver_path:
            options.binary_location = self.driver_path
        if self.profile_path:
            options.add_argument(f"user-data-dir={self.profile_path}")
        if headless:
            options.add_argument("--headless=new")  # for Chrome >= 109
        logger.debug(f"Initializing driver with options: {options}")
        driver = webdriver.Chrome(options=options)
        return driver

    def get_auth_forms(self) -> dict:
        """Gets cookie needed for ms forms api calls"""
        # First, attempt to get cookie from cache

        url = "https://www.office.com/login?es=Click&ru=%2F"
        # For ms forms
        url = "https://go.microsoft.com/fwlink/p/?LinkID=2115709&clcid=0x409&culture=en-us&country=us"
        url = "https://forms.office.com/landing"  # ¿Can be used for checking?
        url = "https://forms.office.com/Pages/DesignPageV2.aspx?origin=Marketing"

        driver = self.get_driver(headless=True)
        driver.get(url)
        driver.implicitly_wait(2)
        ck = driver.get_cookie("OIDCAuth.forms")
        if ck:
            logger.debug("Valid cookie found in cache")
            return ck['value']
        else:
            logger.debug("No valid cooke found. Attempting manually")
        driver.quit()
        driver = self.get_driver(headless=False)
        driver.get(url)
        sessionId = None
        try:
            cookie = WebDriverWait(driver, timeout=120).until(lambda d: d.get_cookie('OIDCAuth.forms'))
            sessionId = cookie['value']
            driver.close()
        finally:
            driver.quit()

        return sessionId

    def get_auth_office(self):
        """Gets token from https://www.office.com"""
        driver = self.get_driver(headless=False)
        # Easier --- office365 main page
        url = "https://www.office.com/login?es=Click&ru=%2F"
        # url = "https://www.office.com/?auth=2"
        # Capture only calls to sharepoint
        driver.scopes = [
            '.*sharepoint.*',
        ]
        driver.get(url)
        token = None
        from time import time
        now = time()
        try:
            req = driver.wait_for_request("sharepoint.com/_api/", timeout=200)
            token = req.headers['Authorization'].split(" ")[-1]
        finally:
            driver.quit()

        return token


if __name__ == '__main__':
    token_manager = SeleniumTokenManager()
    # print(token_manager.get_auth_forms())
    print(token_manager.get_auth_office())


