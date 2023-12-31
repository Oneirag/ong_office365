import unittest

from ong_office365.selenium_token.office365_selenium import SeleniumTokenManager


class TestSeleniumTokenManager(unittest.TestCase):

    def get_tokens(self, clear_cache=False):
        token_manager = SeleniumTokenManager()
        if clear_cache:
            token_manager.clear_cache()

        self.assertIsNotNone(session := token_manager.get_auth_forms_session(),
                             "Could not authenticate interactively for ms forms")
        print(session.cookies)
        self.assertIsNotNone(token := token_manager.get_auth_office(),
                             "Could not authenticate interactively for ms office")
        print(token)

    def test_get_tokens_clear_cache(self):
        self.get_tokens(clear_cache=True)

    def test_get_tokens_with_cache(self):
        self.get_tokens(clear_cache=False)


if __name__ == '__main__':
    unittest.main()
