from __future__ import annotations

import json
import uuid
from abc import abstractmethod
from itertools import count


def serialize_json(dictionary: dict) -> str:
    """Serialize json dictionaries into plain text"""
    return json.dumps(dictionary, ensure_ascii=False, separators=(',', ':'))


class Order:
    def __init__(self):
        self.iterator = count(1000500, 1000000)

    def next(self):
        return next(self.iterator)


class _QuestionInfo:

    def __init__(self, subtitle: str = None, **kwargs):
        self.question = dict()
        if subtitle:
            self.question['ShowSubtitle'] = True
        self.question.update(kwargs)

    def update(self, **kwargs):
        self.question.update(kwargs)

    def payload(self) -> str:
        return serialize_json(self.question)


class _BaseQuestion:
    order = Order()

    @property
    @abstractmethod
    def type(self):
        return None

    def __init__(self, title, is_quiz=False, order: int = None, subtitle: str = None, **kwargs):
        # Create an id based on current timestamp
        new_id = "r" + uuid.uuid4().hex
        self.question = {
            "order": order or self.order.next(),
            "id": new_id,
            "type": self.type,
            "title": title,
            "isQuiz": is_quiz,
        }
        if subtitle:
            self.question['subtitle'] = subtitle
            self.question["formsProRTSubtitle"] = subtitle
        for k, v in kwargs.items():
            if v:
                self.question[k] = v

    def payload(self) -> dict:
        return self.question


class Section(_BaseQuestion):

    @property
    def type(self):
        return "Question.ColumnGroup"

    def __init__(self, title, order: int = None, subtitle: str = None, to_the_end: bool = False):
        if to_the_end:
            # mark this section to send to the end
            question_info = _QuestionInfo(BranchInfo={"ToTheEnd": True}).payload()
        else:
            question_info = None
        super().__init__(title=title, order=order, subtitle=subtitle, questionInfo=question_info)


class QuestionChoice(_BaseQuestion):

    @property
    def type(self):
        return 'Question.Choice'

    def __init__(self, title: str, choices: list | tuple,
                 allow_other_answer: bool = False, subtitle: str = None, required: bool = False,
                 order: int = None):
        """
        Creates a new choice question
        :param title: title of the question
        :param choices: list of strings with the choices, or list of sections to create a menu
        that jumps to each of the sections
        :param allow_other_answer: True to allow multiple options
        :param subtitle: None (default) to leave empty
        :param required: Defaults to False
        :param order: order parameter, to bypass standard ordering
        """
        question_info = _QuestionInfo(subtitle)
        choices_list = []
        for choice in choices:
            append_value = dict()
            if isinstance(choice, str):
                description = choice
                section = None
            else:
                description = choice['title']
                section = choice['id']

            append_value['Description'] = description
            if section:
                append_value['BranchInfo'] = {"TargetQuestionId": section}
                append_value['FormsProDisplayRTText'] = description
            # append_value['IsGenerated']=True
            choices_list.append(append_value)
        question_info.update(Choices=choices_list)
        if not allow_other_answer:
            question_info.update(ChoiceType=1, AllowOtherAnswer=False)

        super().__init__(title=title, questionInfo=question_info.payload(),
                         subtitle=subtitle, required=required, order=order)


class QuestionText(_BaseQuestion):
    @property
    def type(self):
        return "Question.TextField"

    def __init__(self, title: str, multiline: bool = False, subtitle: str = None, order: int = None):
        question_info = _QuestionInfo(subtitle=subtitle, Multiline=multiline)
        super().__init__(title=title, questionInfo=question_info.payload(), subtitle=subtitle, order=order)


if __name__ == '__main__':
    order = Order()

    for _ in range(10):
        print(order.next())

    order = 1
    print(section := Section("Sección").payload())
    order += 1
    print(options1 := QuestionChoice("Pregunta",
                                     choices=['Opción 1', 'Opción 2']).payload())
    print(options2 := QuestionChoice("Pregunta",
                                     choices=['Opción 1', 'Opción 2'], subtitle="Bla bla bla").payload())
    order += 1
    print(text1 := QuestionText("Pregunta larga").payload())
    print(text2 := QuestionText("Pregunta larga", subtitle="bla bla").payload())
    pass
