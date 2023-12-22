"""
Uses the un-documented forms api
"""
from __future__ import annotations

import json
import datetime
import uuid

import pandas as pd
from ong_office365.selenium_token.office365_selenium import SeleniumTokenManager


class Forms:

    def __init__(self):
        self.token_manager = SeleniumTokenManager()
        self.session = self.token_manager.get_auth_forms_session()
        if self.session is None:
            raise ValueError("Could not log in")
        self.__base_url = "https://forms.office.com/formapi/api/"

    def __query_entity(self, entity: str, method="get", json_data=None, params=None):
        return self.__query(method=method, url=self.__base_url + entity, params=params, json_data=json_data)

    def __query(self, url: str, method="get", params=None, json_data=None) -> dict:
        resp = self.session.request(method=method, url=url, params=params, json=json_data)
        resp.raise_for_status()
        try:
            return resp.json()
        except:
            return {'error': resp.content}

    def update_form(self, form_id, **kwargs):
        self.__query_entity(entity=f"forms('{form_id}')", method="patch", json_data=kwargs)

    def get_form_by(self, **kwargs) -> list:
        """Gets list of all forms matching kwargs.
        Example: get_form_by(title='sample title')"""
        filter = [("$filter", f"{key} eq '{value}'") for key, value in kwargs.items()]
        resp = self.__query("https://forms.office.com/formapi/api/forms", params=filter)
        return resp['value']

    def get_forms(self, filter = None) -> list:
        resp = self.__query("https://forms.office.com/formapi/api/forms")
        forms = resp['value']
        retval = []
        for f in forms:
            if filter and filter not in f['title']:
                continue
            retval.append(f)
        return retval

    def get_form_responses(self, form_id: str, all_info: bool=False, question_names: list = None):
        """Use all_info to get a dict with all info about each answer as a dict"""
        resp = self.__query(f"https://forms.office.com/formapi/api/forms('{form_id}')/responses")
        responses = resp['value']
        retval = dict()
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
            if question_names:
                answers_dict.update({question_names[i]: ans['answer1'] for i, ans in enumerate(answers)})
            else:
                answers_dict.update({i: ans['answer1'] for i, ans in enumerate(answers)})

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

    def get_pandas_result(self, form_id: str) -> pd.DataFrame:
        questions = self.get_form_questions(form_id)
        question_names = list(questions.values())
        answers = self.get_form_responses(form_id, all_info=True, question_names=question_names)
        df = pd.DataFrame.from_dict(answers, orient="index")
        # df.columns[:len(names)] = names
        # df.columns = list(questions.values())
        return df

    def create_entity(self, entity: str, **kwargs) -> dict:
        """Entity can be forms, forms('id')/questions, etc, etc"""
        new_entity = self.__query_entity(entity=entity, method="post", json_data=kwargs)
        print(new_entity)
        return new_entity

    def create_form(self, title: str, **kwargs) -> dict:
        return self.create_entity(entity="forms", title=title, **kwargs)


class Question:

    type_option = ""
    def __init__(self, title, **kwargs):
        # Create an id based on current timestamp
        new_id = "r" + uuid.uuid1().hex
        self.question = {
            "title": title,
            "id": new_id
        }
        self.question.update(kwargs)

    @staticmethod
    def serialize_json(dictionary: dict) -> str:
        return json.dumps(dictionary, ensure_ascii=False, separators=(',', ':'))

    def create_option_question(self, choices: list | tuple, allowOtherAnswer: str=False, subtitle: str=None):
        questionInfo = {"Choices": [dict(Description=choice,
                                         # IsGenerated=True
                                         ) for choice in choices],
                        # By default accepts multiple answers. This fixes single answer
                        #,"ChoiceType": 1, "AllowOtherAnswer": False,
                        # "OptionDisplayStyle":"ListAll","ChoiceRestrictionType":"None"
                        }
        if not allowOtherAnswer:
            questionInfo.update({"ChoiceType": 1, "AllowOtherAnswer": False})
        if subtitle:
            questionInfo['ShowSubtitle'] = True

        json_data = {
            'questionInfo': self.serialize_json(questionInfo),
            'type': 'Question.Choice',
            'id': id,
        }
        if subtitle:
            json_data['subtitle'] = subtitle
        json_data.update(self.question)
        return json_data

    def create_text_question(self, multiline:bool=False, subtitle:str=None) -> dict:
        """Use multiline=True for long answers"""
        question_info = dict(MultiLine=multiline)
        if subtitle:
            question_info['ShowSubtitle'] = True
        json_data = {
            "questionInfo": self.serialize_json(question_info),
            "type": "Question.TextField"
        }
        if subtitle:
            json_data['subtitle'] = subtitle
        json_data.update(self.question)
        return json_data

    def create_section(self):
        json_data = {
            "type": "Question.ColumnGroup"
        }
        json_data.update(self.question)
        return json_data


if __name__ == '__main__':
    forms = Forms()
    new_form = forms.create_form("Formulario para borrar")
    form_id = new_form['id']
    print(form_id)
    # my_form = forms.get_form_by(title="Liga de las commodities 2023 - Jornada Diciembre")
    # form_id = my_form[0]['id']
    forms.update_form(form_id, title="Este es otro titulo:" + datetime.datetime.now().isoformat(),
                      description="Subtítulo")

    # To see if sections work
    forms.update_form(form_id, settings='{"ShuffleQuestionOrder":false}')

    # Create an empty session first
    order1 = 2823734420165760
    section = Question("primera sección", order=order1).create_section()
    section_empty = Question("").create_section()
    q_section = forms.create_entity(entity=f"forms('{form_id}')/descriptiveQuestions", **section)
    print(q_section)
    question_text = Question("Pregunta de la seccion1", order=q_section['order']).create_text_question(subtitle="bla bla")
    q_text = forms.create_entity(entity=f"forms('{form_id}')/questions", **question_text)
    print(q_text)

    section = Question("segunda sección", order=order1+1).create_section()
    q_section = forms.create_entity(entity=f"forms('{form_id}')/descriptiveQuestions", **section)
    print(q_section)
    question_choice = Question("Pregunta de ejemplo con opciones", order=q_section['order']).create_option_question(["Una", "Otra", "Y otra"],
                                                                                          subtitle="hola que tal")
    q_choice = forms.create_entity(entity=f"forms('{form_id}')/questions", **question_choice)
    print(q_choice)

    question_text = Question("Pregunta de ejemplo de texto", order=q_section['order']).create_text_question(subtitle="bla bla")
    q_text = forms.create_entity(entity=f"forms('{form_id}')/questions", **question_text)
    print(q_text)


    # question_text = dict(
    #    #  questionInfo=json.dumps(json.dumps(dict(MultiLine=True))),    # False para respuesta corta
    #     title="Pregunta de ejemplo con de respuesta larga",
    #     type="Question.TextField"
    # )
    # q_text = forms.create_entity(entity=f"forms('{form_id}')/questions", **question_text)
    # print(q_text)

    # all_forms = forms.get_forms()
    # print(all_forms)
    # form_id = all_forms[-1]['id']
    my_form = forms.get_form_by(title="One title to search for")
    form_id = my_form[-1]['id']
    df = forms.get_pandas_result(form_id)
    print(df)
    # print(forms.get_form_responses(form_id))
    # print(forms.get_form_questions(form_id))
