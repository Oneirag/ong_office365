"""
Uses the un-documented forms api
"""
import requests
import json
from ong_office365.selenium_token.office365_selenium import SeleniumTokenManager


class Forms:

    def __init__(self):
        self.token_manager = SeleniumTokenManager()
        self.cookie = self.token_manager.get_auth_forms()

    def __query(self, url: str, method="get") -> dict:
        # Token auth does not work!
        # headers = {"Authorization": f"Bearer {token}"}
        headers = {}
        # Cookie from request to https://forms.office.com/Pages/DesignPageV2.aspx?origin=Marketing"
        cookies = {
            'OIDCAuth.forms': self.cookie
        }
        resp = requests.request(method=method, url=url, headers=headers, cookies=cookies)
        resp.raise_for_status()
        try:
            return resp.json()
        except:
            return {'error': resp.content}

    def get_forms(self, filter = None) -> list:
        resp = self.__query("https://forms.office.com/formapi/api/forms")
        forms = resp['value']
        retval = []
        for f in forms:
            if filter and filter not in f['title']:
                continue
            retval.append(f)
        return retval

    def get_form_responses(self, form_id: str):
        resp = self.__query(f"https://forms.office.com/formapi/api/forms('{form_id}')/responses")
        responses = resp['value']
        retval = dict()
        for r in responses:
            answers = json.loads(r['answers'])
            answers_dict = {i: ans['answer1'] for i, ans in enumerate(answers)}
            retval[r['responderName']] = answers_dict
        return retval
        # return responses

    def get_form_questions(self, form_id: str):
        resp = self.__query(f"https://forms.office.com/formapi/api/forms('{form_id}')/questions")
        questions = resp['value']
        retval = dict()
        for q in questions:
            retval[q['id']] = q['title']
            # q['subtitle'] is also interesting
        return retval
        # return responses


if __name__ == '__main__':
    forms = Forms()
    all_forms = forms.get_forms()
    print(all_forms)
    form_id = all_forms[-1]['id']
    print(forms.get_form_responses(form_id))
    print(forms.get_form_questions(form_id))
