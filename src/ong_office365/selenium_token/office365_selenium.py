"""
Authenticate in office365 using selenium
"""

from seleniumwire import webdriver
from selenium.webdriver.support.ui import WebDriverWait

import json
import gzip
from ong_office365 import config, logger


class SeleniumTokenManager:

    def __init__(self):
        self.driver_path = None
        self.profile_path = None
        if "selenium" in config.sections():
            self.driver_path = config["selenium"].get("chrome_driver_path")
            # Path To Custom Profile (needed for using browser cache)
            self.profile_path = config["selenium"].get("profile_path")

    def get_driver(self, headless: bool = False):
        """Initializes a driver, either headless (hidden) or not"""
        options = webdriver.ChromeOptions()
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
        #TODO: Does not work, does not get the API token
        """Gets token from https://www.office.com"""
        driver = webdriver.Chrome()
        # Easier --- office365 main page
        url = "https://www.office.com/login?es=Click&ru=%2F"
        driver.get(url)
        token = None
        try:
            req = driver.wait_for_request("search/_api/SP.OAuth.Token/Acquire",
                                          timeout=200)
            # element = WebDriverWait(driver, 200).until(
            #     # EC.presence_of_element_located((By.ID, "myDynamicElement"))
            #     EC.title_is("SharePoint")
            # )
            token = json.loads(gzip.decompress(req.response.body))
        finally:
            driver.quit()

        driver.close()
        return token


if __name__ == '__main__':
    token_manager = SeleniumTokenManager()
    print(token_manager.get_auth_forms())
    print(token_manager.get_auth_office())