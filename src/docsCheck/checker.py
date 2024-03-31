from datetime import datetime

from docsCheck.utils import *
from math import isclose
import re

import aspose.words as aw


def is_empty_string(string: str):
    empty_symbols = ["\r", "\n", "\r", " ", "\r\n"]
    for char in string:
        if char not in empty_symbols:
            return False

    return True


class UnitChecks:
    doc: aw.Document
    doc_type: str = ""
    doc_type_id: str = ""
    doc_identifier: str = None

    def __init__(self, doc: aw.Document):
        if not type(doc) is aw.Document:
            raise ValueError("doc parameter should provide aspose.words.Document")
        self.doc = doc

    def _check_footers_headers(self, is_header=True):
        """
        :param is_header:
        :param header: False if footer
        :return:
        miss_header_by_section - no header on some page of section
        """

        main_verdict = Verdict(position="Весь документ", standard="19.106-78")
        sections_count = self.doc.sections.count
        has_page_number_by_section = [False] * sections_count
        has_correct_id_by_section = [False] * sections_count
        has_any_header_by_section = [False] * sections_count
        miss_header_by_section = [False] * sections_count

        for i in range(sections_count):
            section = self.doc.sections[i]
            if is_header:
                verdict = Verdict(position=f"Верхний колонтитул раздела {i + 1}", standard="19.106-78")
                headers_array = [section.headers_footers.header_even, section.headers_footers.header_primary]
                if section.page_setup.different_first_page_header_footer:
                    headers_array.append(section.headers_footers.header_first)
            else:
                verdict = Verdict(position=f"Нижний колонтитул раздела {i + 1}", standard="19.106-78")
                headers_array = [section.headers_footers.footer_even, section.headers_footers.footer_primary]
                if section.page_setup.different_first_page_header_footer:
                    headers_array.append(section.headers_footers.header_first)

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

                if header is not None:
                    if not is_empty_string(header.to_string(aw.SaveFormat.TEXT)):
                        section_has_any_header = True
                        if header.is_linked_to_previous:
                            has_page_field = has_page_number_by_section[i - 1]
                            has_correct_id = has_correct_id_by_section[i - 1]
                        else:
                            for field in header.range.fields:
                                if field.as_field().type == aw.fields.FieldType.FIELD_PAGE:
                                    has_page_field = True
                                    break

                            header_text = header.to_string(aw.SaveFormat.TEXT)
                            if is_header:
                                header_text_verdict = self._check_header_text(header_text)
                                has_correct_id = header_text_verdict.ok
                                verdict += header_text_verdict
                            else:
                                verdict += self._check_footer_table(header.tables, header_text)
                                if self._check_identifier(header_text, short=True, exact=False):
                                    verdict += self._check_id_similarity(header_text)
                    else:
                        section_miss_header = True

                    if not has_correct_id:
                        section_has_correct_id = False
                    if not has_page_field:
                        section_has_page_field = False
                    main_verdict += verdict

            has_correct_id_by_section[i] = section_has_correct_id
            has_page_number_by_section[i] = section_has_page_field
            miss_header_by_section[i] = section_miss_header
            has_any_header_by_section[i] = section_has_any_header

        return (
            main_verdict,
            has_correct_id_by_section,
            has_page_number_by_section,
            miss_header_by_section,
            has_any_header_by_section
        )

    def _check_footer_table(self, footer_tables: aw.tables.TableCollection, footer_text) -> Verdict:
        verdict = Verdict(standard="ГОСТ 19.604-78")
        if footer_tables.count > 1:
            verdict.add_message("Нижний колонтитул должен содержать только одну таблицу регистрации изменений.")
        elif footer_tables.count == 1:
            table = footer_tables[0]
            if table.rows.count < 2:
                verdict.add_message("В таблице регистрации изменений недостаточно строк.")

        footer_text = footer_text.lower()
        required = ["изм", "лист", "подп", "дата"]
        contains_required = True
        for word in required:
            if word not in footer_text:
                contains_required = False
                break

        if not (contains_required and re.search(r"№\s+док", footer_text) and re.search("№\s+подл", footer_text)):
            verdict.add_message("В таблице регистрации изменений не присутствуют все необходимые поля.")

        return verdict

    def _check_header_text(self, header_text) -> Verdict:
        verdict = Verdict()
        if self._check_identifier(header_text.strip(), exact=False):
            return self._check_id_similarity(header_text.strip())
        else:
            verdict.ok = False

        return verdict

    @staticmethod
    def _index_paragraph(paragraphs: aw.ParagraphCollection, regexp) -> int:
        proper_tile_index = -1

        for i in range(paragraphs.count):
            node = paragraphs[i]
            paragraph_text = node.as_paragraph().to_string(aw.SaveFormat.TEXT)
            if re.match(regexp, paragraph_text.lower()):
                return i

        return proper_tile_index

    @staticmethod
    def _check_bottom_year(bottom_text: str) -> Verdict:
        verdict = Verdict(ok=True, standard="ГОСТ 19.601-78")
        if re.match(r".*\d{4}.*", bottom_text.strip()):
            if "г" in bottom_text.lower() or "год" in bottom_text.lower():
                verdict.add_message("Строка с указанием года издания (утверждения) не должна содержать 'г' или 'год'.")
        else:
            verdict.add_message(
                "Внизу листа утверждения или титульного листа не содержится указание года издания (утверждения)."
            )

        return verdict

    @staticmethod
    def _check_registration_and_storing(registration_table: aw.tables.Table) -> Verdict:
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

    def _check_identifier(self, identifier: str, short=False, page_type=None, exact=True):
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

        if exact:
            comparing_func = re.match
        else:
            comparing_func = re.search

        if short:
            return comparing_func(r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s", identifier)
        else:
            return comparing_func(
                r"\s*[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d\s" + doc_type_name + r"\s" + doc_type_id + page_type + r"\s*",
                identifier
            )

    def _check_id_similarity(self, identifier: str) -> Verdict:
        verdict = Verdict()
        clean_id = re.search(r"[A-Z]{2}\.\d+\.\d\d\.\d\d-\d\d", identifier)[0]
        if clean_id:
            if self.doc_identifier is None:
                self.doc_identifier = clean_id
            elif self.doc_identifier != clean_id:
                verdict.add_message("Несовпадение идентификатора документа.")

        return verdict

    def _is_text_on_page(self, text, page_number, lower=True) -> bool:
        cloned = self.doc.clone()
        page_text = cloned.extract_pages(int(page_number), 1).to_string(aw.SaveFormat.TEXT)
        if lower:
            page_text = page_text.lower()

        return text in page_text

    def _get_section_page_count(self, section) -> int:
        layout_collector = aw.layout.LayoutCollector(self.doc)

        start_page = layout_collector.get_start_page_index(section)
        end_page = layout_collector.get_end_page_index(section)

        return end_page - start_page + 1

    @staticmethod
    def _find_registration_table(page: aw.Document, verdict: Verdict) -> (bool, Verdict):
        has_registration_table = False
        for node in page.first_section.body.tables:
            table = node.as_table()
            if table.horizontal_anchor == aw.drawing.RelativeHorizontalPosition.PAGE:
                verdict += BaseChecker._check_registration_and_storing(table)
                has_registration_table = True
                break

        return has_registration_table, verdict


class NonTableOfContentsChecker(UnitChecks):
    FLOAT_DELTA: float = 1e-3

    def check_page_margins(self) -> Verdict:
        verdict = Verdict(position="Весь документ", standard="ГОСТ 19.106-78")
        for section in self.doc.sections:
            page_setup = section.as_section().page_setup

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

    def check_lists(self):
        verdict = Verdict(standard="ГОСТ 19.106.78")
        layout_collector = aw.layout.LayoutCollector(self.doc)
        has_hyphen = False
        page_set = set()
        for para in self.doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            paragraph = para.as_paragraph()

            if paragraph.list_format.is_list_item:
                list_level = paragraph.list_format.list_level
                if list_level.number_style == aw.NumberStyle.BULLET:

                    if list_level.number_format == "–" or list_level.number_format == "-":
                        has_hyphen = True
                    else:
                        page_set.add(layout_collector.get_start_page_index(paragraph))

        for page in page_set:
            verdict.add_message(
                "Допускается использовать перечисления только с дефисом.",
                position=f"Страница {page}"
            )

        if has_hyphen:
            verdict.add_message("Рекомендуется использовать только нумерованные перечисления.",
                                position="Весь документ",
                                message_type=MessageTypes.WARNING)

        return verdict

    def check_fonts(self):
        verdict = Verdict()
        right_font = "Times New Roman"

        layout_collector = aw.layout.LayoutCollector(self.doc)
        page_set = set()
        for run in self.doc.get_child_nodes(aw.NodeType.RUN, True):
            # Extract the font name
            font = run.as_run().font
            if font.name != right_font:
                page_set.add(layout_collector.get_start_page_index(run))

        for page_number in page_set:
            verdict.add_message(
                f'Используется некорректный шрифт, используйте "{right_font}" 12 или 14',
                position=f"Страница {page_number}"
            )
        layout_collector.document = None
        return verdict

    def check_line_spacing(self):
        verdict = Verdict(position="Весь документ")

        layout_collector = aw.layout.LayoutCollector(self.doc)
        page_set = set()
        for para in self.doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            paragraph = para.as_paragraph()
            if paragraph.paragraph_format.style.name.startswith("Heading"):
                continue

            line_spacing = paragraph.paragraph_format.line_spacing
            if paragraph.runs[0] is not None:
                if line_spacing != 12 * 1.5:
                    text = para.to_string(aw.SaveFormat.TEXT).strip()
                    if text != "" and not (paragraph.runs[0].font.bold and text.isupper()):
                        page_number = layout_collector.get_start_page_index(para)
                        if 2 < page_number < self.doc.page_count:
                            page_set.add(page_number)

        for page_number in page_set:
            verdict.add_message(
                "Используется некорректный межстрочный интервал",
                position=f"Страница {page_number}"
            )

        return verdict

    def check_headers(self) -> Verdict:
        sections_count = self.doc.sections.count
        (main_verdict, has_correct_id_by_section, has_page_number_by_section, miss_header_by_section,
         has_any_header_by_section) = self._check_footers_headers(is_header=True)

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
                elif has_any_header_by_section[i]:
                    new_verdict = Verdict(position=f"Верхний колонтитул раздела {i + 1}", standard="19.106-78")
                    if not has_page_number_by_section[i]:
                        new_verdict.add_message("Нет номера страницы")
                    if not has_correct_id_by_section[i]:
                        new_verdict.add_message("Некорректный идентификатор документа")

                    main_verdict += new_verdict

            page_count += self._get_section_page_count(self.doc.sections[i])

        return main_verdict

    def check_footers(self):
        sections_count = self.doc.sections.count
        (main_verdict, has_correct_id_by_section, has_page_number_by_section, miss_header_by_section,
         has_any_header_by_section) = self._check_footers_headers(is_header=False)

        page_count = 0
        for i in range(sections_count):
            if page_count < 2:
                if has_any_header_by_section[i]:
                    main_verdict.add_message(
                        "На титульном листе и листе утверждения не должно быть нижнего колонтитула."
                    )
            else:
                if page_count == self.doc.page_count - 1:
                    if has_any_header_by_section[i]:
                        new_verdict = Verdict(position=f"Раздел {i + 1}", standard="19.106-78")
                        new_verdict.add_message("Таблица в нижнем колонтитуле листа регистрации изменений избыточна")
                else:
                    if miss_header_by_section[i]:
                        new_verdict = Verdict(position=f"Раздел {i + 1}", standard="19.106-78")
                        new_verdict.add_message("Пропущен нижний колонтитул в основном тексте документа.")
                        main_verdict += new_verdict
                    elif has_any_header_by_section[i]:
                        new_verdict = Verdict(position=f"Нижний колонтитул раздела {i + 1}", standard="19.106-78")
                        if not has_correct_id_by_section[i]:
                            new_verdict.add_message("Некорректный идентификатор документа")

                        main_verdict += new_verdict

            page_count += self._get_section_page_count(self.doc.sections[i])

        return main_verdict

    def check_certification_page(self) -> Verdict:
        verdict = Verdict(position="Лист утверждения", standard="ГОСТ 19.104-78")
        cloned = self.doc.clone()
        first_page = cloned.extract_pages(0, 1)
        paragraphs = first_page.first_section.body.paragraphs

        proper_tile_index = BaseChecker._index_paragraph(paragraphs, r"\s*лист.+утверждения\s*")

        if proper_tile_index == -1:
            verdict.add_message('Нет надписи "Лист утверждения" на первом листе.')
        elif (proper_tile_index + 1) < paragraphs.count:
            identifier = paragraphs[proper_tile_index + 1].to_string(aw.SaveFormat.TEXT)
            if self._check_identifier(
                    identifier,
                    page_type="ЛУ"
            ):
                verdict += self._check_id_similarity(identifier)
            else:
                verdict.add_message(
                    "Идентификатор документа имеет неверный формат, отсутствует или находится в неположенном месте."
                )

        last_paragraph_text = paragraphs[-1].as_paragraph().to_string(aw.SaveFormat.TEXT)
        verdict += BaseChecker._check_bottom_year(last_paragraph_text)

        has_registration_table, verdict = BaseChecker._find_registration_table(first_page, verdict)
        if not has_registration_table:
            verdict.add_message("Нет таблицы регистрации и хранения или она расположена внутри отступов страницы.")

        return verdict

    def check_title_page(self) -> Verdict:
        verdict = Verdict(position="Титульный лист", standard="ГОСТ 19.104-78")
        title_page = self.doc.extract_pages(1, 1)
        paragraphs = title_page.first_section.body.paragraphs
        proper_tile_index = BaseChecker._index_paragraph(paragraphs, r"\s*листов\s*\d+")

        if proper_tile_index == -1:
            verdict.add_message('Нет надписи о количестве листов на титульном листе.')
        else:
            written_page_count = int(paragraphs[proper_tile_index].to_string(aw.SaveFormat.TEXT).split(" ")[1])
            true_page_count = self.doc.page_count - 1
            if true_page_count != written_page_count:
                verdict.add_message('Некорректное число листов ')

            if (proper_tile_index - 1) >= 0:
                identifier = paragraphs[proper_tile_index - 1].to_string(aw.SaveFormat.TEXT)
                if self._check_identifier(
                        identifier,
                        page_type=None
                ):
                    verdict += self._check_id_similarity(identifier)
                else:
                    verdict.add_message(
                        "Идентификатор документа имеет неверный формат, отсутствует или находится в неположенном месте."
                    )

        has_registration_table, verdict = BaseChecker._find_registration_table(title_page, verdict)
        if not has_registration_table:
            verdict.add_message("Нет таблицы регистрации и хранения или она расположена внутри отступов страницы.")

        first_paragraph_text = paragraphs[0].as_paragraph().to_string(aw.SaveFormat.TEXT)
        if re.match(r"\s*УТВЕРЖД[ЁЕ]Н\s*", first_paragraph_text):
            if 1 < paragraphs.count:
                identifier = paragraphs[1].to_string(aw.SaveFormat.TEXT)
                if self._check_identifier(
                        identifier,
                        page_type="ЛУ"
                ):
                    verdict += self._check_id_similarity(identifier)
                else:
                    verdict.add_message("Отсутствует или некорректен идентификатор листа утверждения")
        else:
            verdict.add_message("Отсутствует пометка об утверждении")

        last_paragraph_text = paragraphs[-1].as_paragraph().to_string(aw.SaveFormat.TEXT)
        verdict += BaseChecker._check_bottom_year(last_paragraph_text)

        return verdict


class BaseChecker(NonTableOfContentsChecker):
    chapters: List[str] = ["аннотация", "содержание", "лист регистрации изменений"]
    doc_standard: str = "ГОСТ 19.106-78"

    toc_valid = False
    names_to_numbers = None
    sorted_numbers = None
    name_to_page = None
    name_to_real_name = None
    name_to_bookmark = None
    has_no_number = None
    numbers_to_names = None

    def main_check(self) -> Verdict:
        main_verdict = Verdict(position="Весь документ.", standard="ГОСТ 19.103-78")

        main_verdict += self.check_page_margins()
        main_verdict += self.check_certification_page()
        main_verdict += self.check_title_page()
        main_verdict += self.check_fonts()

        main_verdict += self.check_footers()
        main_verdict += self.check_headers()
        main_verdict += self.check_table_of_contents()
        # next require table of contents
        main_verdict += self.check_titles()
        main_verdict += self.check_paragraphs()
        main_verdict += self.check_lists()
        main_verdict += self.check_line_spacing()
        main_verdict += self.check_chapters()

        # self.doc.save("WorkingWithComments.add_comments.docx")
        return main_verdict

    def check_chapters(self):
        verdict = Verdict(position="Веcь документ", standard=self.doc_standard)
        if self.toc_valid:
            numerated = set(self.name_to_page.keys())
            if self.toc_valid:
                for chapter in self.chapters:
                    if not (chapter.lower() in self.has_no_number or chapter.lower() in numerated):
                        verdict.add_message(f'Нет необходимого раздела "{chapter}"')

        return verdict

    def check_titles(self) -> Verdict:
        verdict = Verdict(standard="ГОСТ 19.106-78")
        if not self.toc_valid:
            return verdict

        prev_indents_by_level = [0.0] * 4

        for i in range(len(self.sorted_numbers)):
            number = self.sorted_numbers[i]

            title_level = 4 - number.count(0)
            name = self.numbers_to_names[number]
            bookmark = self.name_to_bookmark[name]
            pointer = bookmark.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
            pointed_text = pointer.to_string(aw.SaveFormat.TEXT).strip()
            first_run = pointer.runs[0]
            if first_run:
                next_paragraph = pointer.next_sibling

                # TODO расстояние до предыдущего текста у заголовка подраздела
                while next_paragraph.node_type != aw.NodeType.PARAGRAPH:
                    next_paragraph = next_paragraph.next_sibling
                    if next_paragraph is None:
                        break

                distance_to_next = (pointer.as_paragraph().paragraph_format.space_after
                                    + pointer.as_paragraph().paragraph_format.space_before)

                next_paragraph_text = next_paragraph.to_string(aw.SaveFormat.TEXT).strip()
                has_title_after = False
                if i < len(self.sorted_numbers) - 1:
                    next_title_text = self.name_to_real_name[self.numbers_to_names[self.sorted_numbers[i + 1]]].strip()
                    if next_title_text in next_paragraph_text:
                        has_title_after = True

                if not first_run.font.bold:
                    verdict.add_message(f"Заголовок '{pointed_text}' не выделен жирным шрифтом.")
                if pointed_text[-1] == ".":
                    verdict.add_message(f"Заголовок '{pointed_text}' оканчивается точкой.")

                if title_level == 1:
                    if pointer.paragraph_format.alignment != aw.ParagraphAlignment.CENTER:
                        verdict.add_message(
                            f"Заголовок '{pointed_text}' не центрирован."
                        )
                    if not self.name_to_real_name[name].isupper():
                        verdict.add_message(
                            f"Заголовок уровня 1 '{pointed_text}' написан не строчными буквами."
                        )
                    cloned = self.doc.clone()
                    page = cloned.extract_pages(self.name_to_page[name], 1).sections[0].body

                    is_first = False
                    for node in page.get_child_nodes(aw.NodeType.ANY, True):

                        if node.node_type in [aw.NodeType.PARAGRAPH, aw.NodeType.TABLE]:
                            if node.node_type == aw.NodeType.PARAGRAPH:
                                first_text = node.as_paragraph().to_string(aw.SaveFormat.TEXT).strip()
                                if first_text == self.name_to_real_name[name]:
                                    is_first = True
                            break

                    if not is_first:
                        verdict.add_message(
                            f"Заголовок уровня 1 '{pointed_text}' находится не в начале страницы."
                        )

                    if has_title_after:
                        if distance_to_next < 12 * 3:
                            verdict.add_message(
                                f"Расстояние между заголовком раздела '{pointed_text}' "
                                f"и заголовком подраздела менее, чем 3 высоты шрифта"
                            )

                else:
                    if self.name_to_real_name[name].isupper():
                        verdict.add_message(
                            f"Заголовок уроня {title_level} '{pointed_text}' написан строчными буквами."
                        )

                    left_indent = pointer.paragraph_format.left_indent + pointer.paragraph_format.first_line_indent
                    if title_level != 2 and left_indent <= prev_indents_by_level[title_level - 2]:
                        verdict.add_message(
                            f"Отступ заголовка уровня {title_level} '{pointed_text}' "
                            f"меньше, чем у заголовка предыдущего уровня"
                        )
                    prev_indents_by_level[title_level - 1] = left_indent

                    if not has_title_after:
                        if distance_to_next < 12 * 3 and next_paragraph_text:
                            verdict.add_message(
                                f"Расстояние между заголовком '{pointed_text}' "
                                f"и следующим текстом менее, чем 3 высоты шрифта"
                            )
                            # comment = aw.Comment(self.doc, "Awais Hafeez", "AH", datetime.today())
                            # pointer.append_child(comment)
                            # comment.paragraphs.add(aw.Paragraph(self.doc))
                            # comment.first_paragraph.runs.add(aw.Run(self.doc, f"Расстояние между заголовком "
                            #     f"и следующим текстом менее, чем 3 высоты шрифта"))

        return verdict

    def check_paragraphs(self):
        verdict = Verdict(standard="ГОСТ 19.106.78")
        if not self.toc_valid:
            return verdict

        skip = 2
        cloned = self.doc.clone()
        new_doc = cloned.extract_pages(skip, self.doc.page_count - 3)
        layout_collector = aw.layout.LayoutCollector(new_doc)
        for section in new_doc.sections:
            for node in section.as_section().body.paragraphs:
                paragraph = node.as_paragraph()
                paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT).strip()
                if paragraph_text and not (paragraph_text.lower() in self.has_no_number
                                           or paragraph.paragraph_format.style.name.startswith("TOC")
                                           or re.match(r"(\d+(\.\d+)*\.?\s+)(.*?)$", paragraph_text)
                                           or paragraph.is_list_item):
                    if (paragraph.paragraph_format.first_line_indent <= 0 and
                            paragraph.paragraph_format.alignment != aw.ParagraphAlignment.CENTER):
                        verdict.add_message(
                            f"Абзац текста не имеет абзацного отступа",
                            position=f"Страница {skip + layout_collector.get_start_page_index(paragraph)}")

        return verdict

    def check_table_of_contents(self) -> Verdict:
        verdict = Verdict(position="Содержание", standard="ГОСТ 19.106-78")
        toc_exists = False
        toc_numeration_valid = True
        toc_valid = True
        toc_start = None

        names_to_numbers = {}  # key - name: value - structured number 1.x.x
        numbers_to_names = {}
        unsorted_numbers = []
        name_to_page = {}
        name_to_real_name = {}
        name_to_bookmark = {}
        has_no_number = set()
        was_numerated = False

        allowed_before_numbers = ['аннотация', 'глоссарий']
        allowed_after_numbers = ['лист регистрации изменений']

        for field in self.doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_TOC:
                toc_start = field.start
            if field.type == aw.fields.FieldType.FIELD_HYPERLINK:
                hyperlink = field.as_field_hyperlink()
                if hyperlink.sub_address is not None and hyperlink.sub_address.find("_Toc") == 0:
                    toc_exists = True
                    toc_item = field.start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                    toc_item_text = toc_item.to_string(aw.SaveFormat.TEXT).strip()

                    matched = re.search(r"((\d+(\.\d+)*\.?\s+)|^)(.*?)\s+(\d+)$", toc_item_text)
                    if matched is not None:
                        name_in_toc = matched.group(4)
                        number_in_toc = matched.group(2)
                        page_number = matched.group(5)

                        cleared_name = name_in_toc.strip().lower()
                        name_to_page[cleared_name] = int(page_number)

                        if number_in_toc:
                            was_numerated = True
                            number_in_toc = number_in_toc.strip()
                            if number_in_toc[-1] != ".":
                                verdict.add_message("Номера пунктов должны оканчиваться точкой")
                            number_in_toc = number_in_toc.strip(".")

                            structure_number = list(map(int, number_in_toc.split(".")))
                            if len(structure_number) > 4:
                                toc_numeration_valid = False
                                verdict.add_message(
                                    "Минимальная единица документа - подпункт с номером вида x.x.x.x"
                                    "Более мелкие единицы относятся к перечислениям и в содержании не указываются"
                                )
                                break

                            while len(structure_number) != 4:
                                structure_number.append(0)
                            unsorted_numbers.append(tuple(structure_number))
                            names_to_numbers[cleared_name] = tuple(structure_number)
                            numbers_to_names[tuple(structure_number)] = cleared_name
                        else:
                            has_no_number.add(cleared_name)
                            if was_numerated:
                                if not (cleared_name in allowed_after_numbers or "приложение" in cleared_name):
                                    verdict.add_message(
                                        f"Пункт {name_in_toc} должен быть пронумерован "
                                        f"или находиться перед содержанием документа"
                                    )
                            else:
                                if not (cleared_name in allowed_before_numbers):
                                    verdict.add_message(
                                        f"Пункт {name_in_toc} должен быть пронумерован или находиться в конце документа"
                                    )

                        bookmark = self.doc.range.bookmarks.get_by_name(hyperlink.sub_address)
                        try:
                            name_to_bookmark[name_in_toc.lower().strip()] = bookmark
                            pointer = bookmark.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                            pointed_text = pointer.to_string(aw.SaveFormat.TEXT)
                            matched = re.search(r"((\d+(\.\d+)*\.?\s+)|^)(.*?)$", pointed_text)
                            real_name = matched.group(4).strip()
                            if not (cleared_name == pointed_text.lower().strip() or real_name.lower() == cleared_name):
                                verdict.add_message(
                                    f'Заголовок содержания "{name_in_toc}" не совпадает с заголовком в тексте'
                                )
                                toc_valid = False
                            else:
                                name_to_real_name[cleared_name] = real_name

                        except Exception:
                            if not self._is_text_on_page(name_in_toc.lower().strip(), page_number):
                                verdict.add_message(
                                    f'Заголовок содержания "{name_in_toc}" не совпадает с заголовком в тексте'
                                )
                                toc_valid = False
                            else:
                                name_to_real_name[cleared_name] = name_in_toc.strip()

        if not toc_exists:
            verdict.add_message("В документе нет содержания")
            return verdict

        layout_collector = aw.layout.LayoutCollector(self.doc)
        toc_start_page = layout_collector.get_start_page_index(toc_start)  # this is 1-based index
        name_to_page["содержание"] = toc_start_page

        if not self._is_text_on_page("СОДЕРЖАНИЕ", toc_start_page - 1, lower=False):
            verdict.add_message("Страница содержания должна содержать заголовок 'СОДЕРЖАНИЕ'")

        sorted_numbers = sorted(unsorted_numbers)
        if toc_numeration_valid:
            if sorted_numbers != unsorted_numbers:
                verdict.add_message("Нарушен порядок нумерации в содержании и тексте документа")
            if sorted_numbers:
                first_element_page = name_to_page[numbers_to_names[sorted_numbers[0]]]
                if first_element_page <= toc_start_page:
                    verdict.add_message("Содержание должно находиться перед основным текстом на отдельной странице.")

        if "аннотация" in has_no_number:
            if toc_start_page < name_to_page["аннотация"]:
                verdict.add_message("Аннотация должна быть перед содержанием")
        else:
            verdict.add_message("Аннотация не нумеруется")

        if "содержание" in has_no_number:
            verdict.add_message("Содержание не указывается в содержании и не нумеруется.")

        if "лист регистрации изменений" not in has_no_number:
            verdict.add_message("Лист регистрации изменений не нумеруется")

        if toc_valid and toc_numeration_valid:
            self.toc_valid = True
            self.names_to_numbers = names_to_numbers
            self.sorted_numbers = sorted_numbers
            self.name_to_page = name_to_page
            self.name_to_real_name = name_to_real_name
            self.name_to_bookmark = name_to_bookmark
            self.has_no_number = has_no_number
            self.numbers_to_names = numbers_to_names

        return verdict


class TechTaskChecker(BaseChecker):
    doc_type: str = "ТЗ"
    doc_type_id: str = "01-1"
    chapters: List[str] = [
        "аннотация", "введение", "содержание",
        "лист регистрации изменений", "назначение разработки",
        "требования к программе",
        "требования к программной документации",
        "технико-экономические показатели",
        "стадии и этапы разработки",
        "порядок контроля и приемки"
    ]
    doc_standard: str = "ГОСТ 19.106-78"


class OperatorManualChecker(BaseChecker):
    doc_type: str = "34"
    doc_type_id: str = "01-1"
    chapters: List[str] = [
        "содержание",
        "лист регистрации изменений",
        "назначение программы",
        "условия выполнения программы",
        "выполнение программы",
        "сообщения автору"
    ]
    doc_standard: str = "ГОСТ 19.505-79"


class ExplanatoryNoteChecker(BaseChecker):
    doc_type: str = "81"
    doc_type_id: str = "01-1"
    chapters: List[str] = [
        "содержание",
        "лист регистрации изменений",
        "введение",
        "назначение и область применения",
        "технические характеристики",
        "ожидаемые технико-экономические показатели",
        "источники, использованные при разработке"
    ]
    doc_standard: str = "ГОСТ 19.404-79"


allowed_checkers = {
    "ОБЩЕЕ": BaseChecker,
    "ТЗ": TechTaskChecker,
    "РО": OperatorManualChecker,
    "ПЗ": ExplanatoryNoteChecker,
}

full_allowed_checkers_name = {
    "ОБЩЕЕ": "Только общая проверка",
    "ТЗ": "Техническое задание",
    "РО": "Руководство оператора",
    "ПЗ": "Пояснительная записка"
}
