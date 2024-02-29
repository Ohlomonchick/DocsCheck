from utils import *
from math import isclose
import re

import aspose.words as aw


def is_empty_string(string: str):
    empty_symbols = ["\r", "\n", "\r", " ", "\r\n"]
    for char in string:
        if char not in empty_symbols:
            return False

    return True


class BaseChecker:
    doc: aw.Document
    FLOAT_DELTA: float = 1e-3
    doc_type: str = "ТЗ"
    doc_type_id: str = "01-1"
    doc_identifier: str = None

    def __init__(self, doc: aw.Document):
        if not type(doc) is aw.Document:
            raise ValueError("doc parameter should provide aspose.words.Document")
        self.doc = doc

    def check_page_margins(self) -> Verdict:
        verdict = Verdict(position="Весь документ", standard="ГОСТ 19.106-78")
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

        footer = self.doc.sections[0].headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]
        footer.paragraphs.add_run("Footer Text")
        pass

    def check_headers(self) -> Verdict:
        main_verdict = Verdict(position="Весь документ", standard="19.106-78")
        sections_count = self.doc.sections.count
        has_page_number_by_section = [False] * sections_count
        has_correct_id_by_section = [False] * sections_count
        has_any_header_by_section = [False] * sections_count
        miss_header_by_section = [False] * sections_count

        for i in range(sections_count):
            verdict = Verdict(position=f"Верхний колонтитул раздела {i + 1}", standard="19.106-78")
            section = self.doc.sections[i]

            headers_array = [
                section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY]
            ]

            if section.page_setup.odd_and_even_pages_header_footer:
                headers_array.append(section.headers_footers[aw.HeaderFooterType.HEADER_EVEN])
            if section.page_setup.different_first_page_header_footer:
                headers_array.append(section.headers_footers[aw.HeaderFooterType.HEADER_FIRST])

            if not any(headers_array):
                # linked to previous whole
                has_correct_id_by_section[i] = has_correct_id_by_section[i - 1]
                has_page_number_by_section[i] = has_page_number_by_section[i - 1]
                miss_header_by_section[i] = miss_header_by_section[i - 1]
                has_any_header_by_section[i] = has_any_header_by_section[i - 1]
                continue

            section_miss_header = False
            section_has_any_header = False
            section_has_page_field = True
            section_has_correct_id = True
            for header in headers_array:
                has_page_field = False
                has_correct_id = True
                if header is not None and not is_empty_string(header.to_string(aw.SaveFormat.TEXT)):
                    section_has_any_header = True
                    if header.is_linked_to_previous:
                        has_page_field = has_page_number_by_section[i - 1]
                        has_correct_id = has_correct_id_by_section[i - 1]
                    else:
                        print(header.to_string(aw.SaveFormat.TEXT))
                        for field in header.range.fields:
                            if field.type == aw.fields.FieldType.FIELD_PAGE:
                                has_page_field = True
                                break

                        header_text = header.to_string(aw.SaveFormat.TEXT)
                        if self.check_identifier(header_text.strip()):
                            if self.check_id_similarity(header_text.strip()):
                                verdict.add_message("Несовпадение идентификатора документа.")
                                has_correct_id = False
                        else:
                            has_correct_id = False
                else:
                    section_miss_header = True

                if not has_correct_id and section_has_correct_id:
                    verdict.add_message("Некорректный идентификатор документа.")
                    section_has_correct_id = False
                if not has_page_field and section_has_page_field:
                    section_has_page_field = False
                    verdict.add_message("Отсутствует нумерация страниц.")

            has_correct_id_by_section[i] = section_has_correct_id
            has_page_number_by_section[i] = section_has_page_field
            miss_header_by_section[i] = section_miss_header
            has_any_header_by_section[i] = section_has_any_header

        page_count = 0
        for i in range(sections_count):
            if page_count < 2:
                if has_any_header_by_section[i]:
                    main_verdict.add_message(
                        "На титульном листе и листе утверждения не должно быть верхнего колонтитула."
                    )
            else:
                if miss_header_by_section[i]:
                    new_verdict = Verdict(position=f"Раздел {i + 1}", standard="19.106-78")
                    new_verdict.add_message("Пропущен верхний колонтитул в основном тексте документа.")
                    main_verdict += new_verdict

            page_count += self.get_section_page_count(self.doc.sections[i])

        return main_verdict

    def get_section_page_count(self, section):
        layout_collector = aw.layout.LayoutCollector(self.doc)

        start_page = layout_collector.get_start_page_index(section)
        end_page = layout_collector.get_end_page_index(section)

        return end_page - start_page + 1

    def check_certification_page(self) -> Verdict:
        verdict = Verdict(position="Лист утверждения", standard="ГОСТ 19.104-78")
        first_page = self.doc.extract_pages(0, 1)
        paragraphs = first_page.first_section.body.paragraphs

        proper_tile_index = BaseChecker.index_paragraph(paragraphs, r"\s*лист.+утверждения\s*")

        if proper_tile_index == -1:
            verdict.add_message('Нет надписи "Лист утверждения" на первом листе.')
        elif (proper_tile_index + 1) < paragraphs.count:
            identifier = paragraphs[proper_tile_index + 1].to_string(aw.SaveFormat.TEXT)
            if self.check_identifier(
                    identifier,
                    page_type="ЛУ"
            ):
                verdict += self.check_id_similarity(identifier)
            else:
                verdict.add_message(
                    "Идентификатор документа имеет неверный формат, отсутствует или находится в неположенном месте."
                )

        last_paragraph_text = paragraphs[-1].as_paragraph().to_string(aw.SaveFormat.TEXT)
        verdict += BaseChecker.check_bottom_year(last_paragraph_text)

        has_registration_table, verdict = BaseChecker._find_registration_table(first_page, verdict)
        if not has_registration_table:
            verdict.add_message("Нет таблицы регистрации и хранения или она расположена внутри отступов страницы.")

        return verdict

    @staticmethod
    def _find_registration_table(page: aw.Document, verdict: Verdict) -> (bool, Verdict):
        has_registration_table = False
        for node in page.first_section.body.tables:
            table = node.as_table()
            if table.horizontal_anchor == aw.drawing.RelativeHorizontalPosition.PAGE:
                verdict += BaseChecker.check_registration_and_storing(table)
                has_registration_table = True
                break

        return has_registration_table, verdict

    def check_title_page(self) -> Verdict:
        verdict = Verdict(position="Титульный лист", standard="ГОСТ 19.104-78")
        title_page = self.doc.extract_pages(1, 1)
        paragraphs = title_page.first_section.body.paragraphs

        proper_tile_index = BaseChecker.index_paragraph(paragraphs, r"\s*листов\s*\d+")

        if proper_tile_index == -1:
            verdict.add_message('Нет надписи о количестве листов на титульном листе.')
        else:
            written_page_count = int(paragraphs[proper_tile_index].to_string(aw.SaveFormat.TEXT).split(" ")[1])
            true_page_count = self.doc.page_count - 1
            if true_page_count != written_page_count:
                verdict.add_message('Некорректное число листов ')

            if (proper_tile_index - 1) >= 0:
                identifier = paragraphs[proper_tile_index - 1].to_string(aw.SaveFormat.TEXT)
                if self.check_identifier(
                        identifier,
                        page_type=None
                ):
                    verdict += self.check_id_similarity(identifier)
                else:
                    verdict.add_message(
                        "Идентификатор документа имеет неверный формат, отсутствует или находится в неположенном месте."
                    )

        has_registration_table, verdict = BaseChecker._find_registration_table(title_page, verdict)
        if not has_registration_table:
            verdict.add_message("Нет таблицы регистрации и хранения или она расположена внутри отступов страницы.")

        first_paragraph_text = paragraphs[0].as_paragraph().to_string(aw.SaveFormat.TEXT)
        if re.match("\s*УТВЕРЖД(Ё|Е)Н\s*", first_paragraph_text):
            if 1 < paragraphs.count:
                identifier = paragraphs[1].to_string(aw.SaveFormat.TEXT)
                if self.check_identifier(
                        identifier,
                        page_type="ЛУ"
                ):
                    verdict += self.check_id_similarity(identifier)
                else:
                    verdict.add_message("Отсутствует или некорректен идентификатор листа утверждения")
        else:
            verdict.add_message("Отсутствует пометка об утверждении")

        last_paragraph_text = paragraphs[-1].as_paragraph().to_string(aw.SaveFormat.TEXT)
        verdict += BaseChecker.check_bottom_year(last_paragraph_text)

        return verdict

    @staticmethod
    def index_paragraph(paragraphs: aw.ParagraphCollection, regexp):
        proper_tile_index = -1

        for i in range(paragraphs.count):
            node = paragraphs[i]
            paragraph_text = node.as_paragraph().to_string(aw.SaveFormat.TEXT)
            if re.match(regexp, paragraph_text.lower()):
                return i

        return proper_tile_index

    @staticmethod
    def check_bottom_year(bottom_text: str) -> Verdict:
        verdict = Verdict(ok=True, standard="ГОСТ.601-78")
        if re.match(r".*\d{4}.*", bottom_text.strip()):
            if "г" in bottom_text.lower() or "год" in bottom_text.lower():
                verdict.add_message("Строка с указанием года издания (утверждения) не должна содержать 'г' или 'год'.")
        else:
            verdict.add_message(
                "Внизу листа утверждения или титульного листа не содержится указание года издания (утверждения)."
            )

        return verdict

    @staticmethod
    def check_registration_and_storing(registration_table: aw.tables.Table) -> Verdict:
        verdict = Verdict(standard="ГОСТ.601-78")
        if registration_table.rows.count != 5:
            verdict.add_message("В таблице регистрации и хранения должно быть 5 колонок.")
            return verdict

        left_length = registration_table.absolute_horizontal_distance
        for cell in registration_table.rows[0].as_row().cells:
            left_length += cell.as_cell().cell_format.width

        if left_length > aw.ConvertUtil.millimeter_to_point(20):
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

        if self.doc_type_id is not None:
            doc_type_id = self.doc_type_id

        if short:
            return re.match(r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s", identifier)
        else:
            return re.match(
                r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s" + doc_type_name + r"\s" + doc_type_id + page_type + r"\s*",
                identifier
            )

    def check_id_similarity(self, identifier: str) -> Verdict:
        verdict = Verdict()
        clean_id = re.search(r"[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d", identifier)[0]
        if clean_id:
            if self.doc_identifier is None:
                self.doc_identifier = clean_id
            elif self.doc_identifier != clean_id:
                verdict.add_message("Несовпадение идентификатора документа.")

        return verdict
