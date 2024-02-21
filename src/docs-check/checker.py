from dataclasses import dataclass
from typing import List
from math import isclose
import re

import aspose.words as aw


@dataclass
class Message:
    text: str
    position: str
    standard: str


class Verdict:
    """Class for storing result"""
    ok: bool
    messages: List[Message]
    standard: str
    position: str

    def __init__(self, ok: bool = True, messages: List[Message] = None, position: str = None, standard: str = None):
        self.ok = ok
        self.messages = messages
        self.position = position
        self.standard = standard

        if messages is None:
            self.messages = []
        if position is None:
            self.position = ""
        if standard is None:
            self.standard = ""

    def add_message(self, message: str):
        self.messages.append(Message(message, position=self.position, standard=self.standard))
        self.ok = False

    def __add__(self, other):
        self.messages += other.messages
        if not other.ok:
            self.ok = False
        return self


class BaseChecker:
    doc: aw.Document
    FLOAT_DELTA: float = 1e-3
    doc_type: str = None
    doc_identifier: str = None

    def __init__(self, doc: aw.Document):
        if not type(doc) is aw.Document:
            raise ValueError("doc parameter should provide aspose.words.Document")
        self.doc = doc

    def check_page_margins(self) -> Verdict:
        """ГОСТ 19.106-78"""
        verdict = Verdict()
        page_setup = self.doc.sections[0].page_setup

        if page_setup.orientation != aw.Orientation.PORTRAIT:
            verdict.add_message("Некорректная ориентация страницы. Она должна быть книжной.")
        if page_setup.paper_size != aw.PaperSize.A4:
            verdict.add_message(
                f"Документация оформляется на листах формата А4. Ваш формат - {page_setup.paper_size}"
            )
        if not isclose(page_setup.left_margin, aw.ConvertUtil.millimeter_to_point(20), rel_tol=self.FLOAT_DELTA):
            verdict.add_message(
                f"Неверный отступ слева. Требуемый - 20мм."
            )
        if not isclose(page_setup.right_margin, aw.ConvertUtil.millimeter_to_point(10), rel_tol=self.FLOAT_DELTA):
            verdict.add_message(
                f"Неверный отступ справа. Требуемый - 10мм."
            )
        if not isclose(page_setup.bottom_margin, aw.ConvertUtil.millimeter_to_point(15), rel_tol=self.FLOAT_DELTA):
            verdict.add_message(
                f"Неверный отступ снизу. Требуемый - 15мм."
            )
        if not isclose(page_setup.top_margin, aw.ConvertUtil.millimeter_to_point(25), rel_tol=self.FLOAT_DELTA):
            verdict.add_message(
                f"Неверный отступ сверху. Требуемый - 25мм."
            )

        return verdict

    def check_changes_string(self) -> Verdict:
        # TODO
        header = self.doc.sections[0].headers_footers[aw.HeaderFooterType.HEADER_PRIMARY]
        header.paragraphs.add_run("Header Text")

        footer = self.sections[0].headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]
        footer.paragraphs.add_run("Footer Text")
        pass

    def check_certification_page(self) -> Verdict:
        """ГОСТ 19.104-78"""
        verdict = Verdict()
        first_page = self.doc.extract_pages(0, 1)
        paragraphs = first_page.first_section.body.paragraphs

        proper_tile_index = -1
        for i in range(paragraphs.count):
            node = paragraphs[i]
            paragraph_text = node.as_paragraph().to_string(aw.SaveFormat.TEXT)
            if re.match(r"\s*лист.+утверждения\s*", paragraph_text.lower()):
                proper_tile_index = i
                break

        if proper_tile_index == -1:
            verdict.add_message('Нет надписи "Лист утверждения" на титульном листе.')
        elif (proper_tile_index + 1) < paragraphs.count:
            if not self.check_identifier(
                    paragraphs[proper_tile_index + 1].to_string(aw.SaveFormat.TEXT),
                    page_type="ЛУ"
            ):
                verdict.add_message("Идентификатор документа имеет неверный формат или отсутствует.")

        last_paragraph_text = paragraphs[-1].as_paragraph().to_string(aw.SaveFormat.TEXT)
        verdict += BaseChecker.check_bottom_year(last_paragraph_text)

        has_registration_table = False
        for node in first_page.first_section.body.tables:
            table = node.as_table()
            if table.horizontal_anchor == aw.drawing.RelativeHorizontalPosition.PAGE:
                verdict += BaseChecker.check_registration_and_storing(table)
                has_registration_table = True
                break
        if not has_registration_table:
            verdict.add_message("Нет таблицы регистрации и хранения или она расположена внутри отступов страницы.")

        return verdict

    @staticmethod
    def check_bottom_year(bottom_text: str) -> Verdict:
        verdict = Verdict(ok=True)
        if re.match(r".*\d{4}.*", bottom_text):
            if "г" in bottom_text.lower() or "год" in bottom_text.lower():
                verdict.add_message("Строка с указанием года издания (утверждения) не должна содержать 'г' или 'год'.")
        else:
            verdict.add_message(
                "Внизу листа утверждения или титульного листа не содержится указание года издания (утверждения)."
            )

        return verdict
    @staticmethod
    def check_registration_and_storing(registration_table: aw.tables.Table) -> Verdict:
        """ГОСТ.601-78"""
        verdict = Verdict(standard="ГОСТ.601-78")
        if registration_table.rows.count != 5:
            verdict.add_message("В таблице регистрации и хранения должно быть 5 колонок.")
            return verdict

        left_length = registration_table.absolute_horizontal_distance
        for cell in registration_table.rows[0].as_row().cells:
            left_length += cell.as_cell().cell_format.width

        if left_length > aw.ConvertUtil.millimeter_to_point(25):
            verdict.add_message(
                f"Таблица регистрации и хранения должна быть за левым полем документа полностью."
            )

        whole_text = ""
        for node in registration_table.rows:
            whole_text += node.as_row().cells[0].as_cell().get_text()
        whole_text = whole_text.lower()
        if not ("инв" in whole_text and "под" in whole_text and "дата" in whole_text):
            verdict.add_message("В таблице регистрации и хранения отсутствуют необходимые надписи.")

        return verdict

    def check_identifier(self, identifier: str, short=False, page_type=None):
        if page_type is None:
            page_type = ""
        else:
            page_type = "-" + page_type

        doc_type_name = r"\w\w"
        doc_type_id = r"\d\d-\d"
        if self.doc_type is not None:
            doc_type_name = self.doc_type

        if self.doc_identifier is not None:
            doc_type_id = self.doc_identifier

        if short:
            return re.match(r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s", identifier)
        else:
            return re.match(
                r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s" + doc_type_name + r"\s" + doc_type_id + page_type + r"\s*",
                identifier
            )

