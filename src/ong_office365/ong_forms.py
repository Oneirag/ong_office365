"""
Uses the un-documented forms api
"""
from __future__ import annotations

import datetime
import json

import pandas as pd

from ong_office365 import logger as log
from ong_office365.forms_objects.questions import Section, QuestionText, QuestionChoice
from ong_office365.selenium_token.office365_selenium import SeleniumTokenManager


def remove_sections(questions: list) -> list:
    """Removes sections from question list"""
    retval = list(q for q in questions if q['type'] not in ('Question.ColumnGroup',))
    return retval


class Forms:

    def __init__(self, logger=None):
        self.logger = logger or log
        self.token_manager = SeleniumTokenManager(logger=logger)
        self.session = None
        self.login()
        if self.session is None:
            raise ValueError("Could not log in")
        self.__base_url = "https://forms.office.com"
        self.__api_base_url = f"{self.__base_url}/formapi/api/"

    def login(self, fresh=False):
        if fresh:
            self.token_manager.clear_cache()
        self.token_manager.get_auth_forms_session()
        self.session = self.token_manager.get_auth_forms_session()

    def __query_entity(self, entity: str, method="get", json_data=None, params=None):
        return self.__query(method=method, url=self.__api_base_url + entity, params=params, json_data=json_data)

    def __query(self, url: str, method="get", params=None, json_data=None, retry=False) -> dict:
        resp = self.session.request(method=method, url=url, params=params, json=json_data)
        try:
            resp.raise_for_status()
        except:
            if not retry:
                if "AntiForgery token validation error" in resp.text:
                    # Try again with a new fresh login
                    self.login(fresh=True)
                    return self.__query(url, method, params, json_data, retry=True)
            else:
                print(resp.content)
                raise
        try:
            return resp.json()
        except:
            return {'error': resp.content}

    def update_form(self, form_id, **kwargs):
        self.__query_entity(entity=f"forms('{form_id}')", method="patch", json_data=kwargs)

    def get_form_by(self, **kwargs) -> list:
        """Gets list of all forms matching kwargs. By default, it does not return softDeleted forms.
        Example: get_form_by(title='sample title')
        If you want to show all forms (including deleted), use get_form_by(title="sample title", softDeleted=None)
        If you want to show deleted forms, use get_form_by(title="sample title", softDeleted=1)
        """
        softDeleted = kwargs.pop("softDeleted", 0)
        filter = [("$filter", f"{key} eq '{value}'") for key, value in kwargs.items()]
        resp = self.__query_entity("forms", params=filter)
        forms = resp['value']
        # Removes softDeleted manually...as filter does not work!
        if softDeleted is not None:
            forms = [f for f in forms if f['softDeleted'] == int(softDeleted)]

        return forms

    def get_forms(self, filter=None) -> list:
        resp = self.__query_entity("forms")
        forms = resp['value']
        retval = []
        for f in forms:
            if filter and filter not in f['title']:
                continue
            retval.append(f)
        return retval

    def get_form_responses(self, form_id: str, all_info: bool = False) -> list:
        """Uses all_info to get a dict with all info about each answer as a list. It can return multiple answers for the
        same user"""
        questions = self.get_form_questions(form_id)
        resp = self.__query_entity(f"forms('{form_id}')/responses")
        responses = resp['value']
        retval = list()
        for idx, r in enumerate(responses):
            answers = json.loads(r['answers'])
            answers_dict = dict()
            # Add ID	Hora de inicio	Hora de finalización	Correo electrónico	Nombre
            if all_info:
                answers_dict['ID'] = idx + 1
                answers_dict['Hora de inicio'] = pd.Timestamp(r['startDate'])
                answers_dict['Hora de finalizacion'] = pd.Timestamp(r['submitDate'])
                answers_dict['Correo electrónico'] = r['responder']
                answers_dict['Nombre'] = r['responderName']
            for ans in answers:
                question_id = ans['questionId']
                if question_id in questions:
                    answers_dict[questions[question_id]] = ans['answer1']
            # retval[r['responderName']] = answers_dict
            retval.append(answers_dict)
        return retval

    def get_form_questions(self, form_id: str):
        resp = self.__query_entity(f"forms('{form_id}')/questions")
        # Removes sections
        questions = remove_sections(resp['value'])
        retval = dict()
        for q in questions:
            retval[q['id']] = q['title']
            # q['subtitle'] is also interesting
        return retval

    def get_pandas_result(self, form_id: str) -> pd.DataFrame:
        answers = self.get_form_responses(form_id, all_info=True)
        df = pd.DataFrame(answers).set_index("Nombre", drop=False)
        return df

    def create_entity(self, entity: str, **kwargs) -> dict:
        """Entity can be forms, forms('id')/questions, etc, etc"""
        new_entity = self.__query_entity(entity=entity, method="post", json_data=kwargs)
        self.logger.debug(new_entity)
        return new_entity

    def create_form(self, title: str, settings='{"ShuffleQuestionOrder":false}', **kwargs) -> dict:
        """
        Creates a new form instance and returns json object created
        :param title: title (name) of the form
        :param settings: defaults to settings='{"ShuffleQuestionOrder":false}', needed for forms with sections
        :param kwargs: additional key/value pairs to send to server. Example: description="subtitle"
        :return: a json object
        """
        return self.create_entity(entity="forms", title=title, settings=settings, **kwargs)

    def create_section(self, form_id, section: Section) -> dict:
        self.logger.debug(section.payload())
        q_section = self.create_entity(entity=f"forms('{form_id}')/descriptiveQuestions", **section.payload())
        return q_section

    def create_question(self, form_id, question: QuestionChoice | QuestionText):
        self.logger.debug(question.payload())
        q_question = self.create_entity(entity=f"forms('{form_id}')/questions", **question.payload())
        return q_question

    def delete_form(self, form_id: str):
        """Permanently deletes form"""
        self.__query_entity(f"forms('{form_id}')", method="delete")

    def trash_form(self, form_id: str):
        """Sends form to trash bin"""
        self.__query_entity(f"forms('{form_id}')", method="patch",
                            json_data={"softDeleted": 1, "collectionId": None})

    def get_form_response_page_url(self, form_id: str) -> str:
        """Returns url to respond to a form"""
        return f"{self.__base_url}/Pages/ResponsePage.aspx?id={form_id}"


if __name__ == '__main__':
    forms = Forms()
    exit(0)

    new_form = forms.create_form("Deletable form: " + datetime.datetime.now().isoformat(),
                                 description="Subtitle goes here",
                                 # These settings are added when creating sections
                                 settings='{"ShuffleQuestionOrder":false}',
                                 )
    form_id = new_form['id']
    print(form_id)
    # forms.update_form(form_id, title="Another title:" + datetime.datetime.now().isoformat(),
    #                   description="Subtitle")
    # forms.update_form(form_id, description="Subtitle goes here",
    #                  settings='{"ShuffleQuestionOrder":false}')

    # Create a first section, where menu will be hosted
    section = Section(title="Main menu", subtitle="Choose section")
    menu_section = forms.create_section(form_id, section)
    # Create sections
    sections = []
    for section_id in range(3):
        # Creates a new section that jumps directly to the end after filling it
        section = Section(title=f"Section {section_id}", subtitle=f"Sample section {section_id}",
                          to_the_end=True)
        q_section = forms.create_section(form_id, section)
        sections.append(q_section)
        # Add some questions to the section
        question_text = QuestionText(title=f"Long text question of section {section_id}", multiline=True)
        q_text = forms.create_question(form_id, question_text)
        question_option = QuestionChoice(title=f"Choice question of section {section_id}",
                                         choices=["One", "Two"], subtitle="Select one")
        q_option = forms.create_question(form_id, question_option)
        question_text = QuestionText(title=f"Short text question for section {section_id}", multiline=False,
                                     subtitle="Just one line")
        q_text = forms.create_question(form_id, question_text)
        question_option = QuestionChoice(title=f"Multiple choice question of section {section_id}",
                                         choices=["One", "Two", "three"], subtitle="Select multiple",
                                         allow_other_answer=True)
        q_option = forms.create_question(form_id, question_option)

    # Now, create a question menu that sends to each of the branches
    menu = QuestionChoice(title="", choices=sections, order=menu_section['order'] + 1)
    q_menu = forms.create_question(form_id, menu)

    # all_forms = forms.get_forms()
    # print(all_forms)
    # form_id = all_forms[-1]['id']
    title_to_search = "One title to search for"
    my_form = forms.get_form_by(title="One title to search for")
    if my_form:
        form_id = my_form[-1]['id']
        df = forms.get_pandas_result(form_id)
        print(df)
    else:
        print(f"Form '{title_to_search}' not found")
