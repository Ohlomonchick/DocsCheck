import os
import sys
from docsCheck import runners


HELP = """ИСПОЛЬЗОВАНИЕ:
docsCheck <path_to_docx> <doc_type>

где:
path_to_docx - абсолютный или относительный путь до анализируемого файла
doc_type - один из доступных типов документов (опционально)

Доступные типы документов:
ТЗ - техническое задание
"""


def print_verdict(verdict):
    for message in verdict.messages:
        print(message)


def main():
    args = sys.argv[1:]
    if len(args) > 2 or len(args) < 1:
        print("Неверное количество аргументов!")
        print(HELP)
        return

    if args[0] == "--help":
        print(HELP)
        return

    workdir_path = os.getcwd()
    if os.path.isabs(args[0]):
        doc_path = args[0]
    else:
        doc_path = os.path.join(workdir_path, args[0])

    if not os.path.isfile(doc_path):
        print(f"Путь {doc_path} не является файлом")
        return

    filename, extension = os.path.splitext(doc_path)
    if extension != ".docx":
        print("Файл должен иметь расширение docx")
        return

    allowed_doc_types = ["ТЗ"]

    doc_type = None
    if len(args) == 2:
        if args[1] in allowed_doc_types:
            doc_type = args[1]
        else:
            print(f"Тип документа {args[1]} недоступен")
            print(HELP)
            return

    verdict = runners.run_check(doc_path, doc_type)
    print_verdict(verdict)
