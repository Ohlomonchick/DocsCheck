import aspose.words as aw
import pathlib
import os
from docsCheck import checker


def run_check(doc_path, doc_type=None, licence_path=None):
    if licence_path is None:
        package_path = pathlib.Path(__file__).parent.resolve()
        licence_path = os.path.join(package_path, "Aspose.WordsforPythonvia.NET.lic")

    lic = aw.License()
    try:
        lic.set_license(licence_path)
    except RuntimeError as err:
        print("\nThere was an error setting the license:", err)
        return
    try:
        doc = aw.Document(doc_path)
    except RuntimeError:
        print("Невозможно открыть документ. Возможно, он используется другим процессом")
        return
    except Exception:
        print("Файл повреждён.")
        return

    try:
        if doc_type is None:
            check = checker.BaseChecker(doc)
        else:
            check = checker.allowed_checkers[doc_type](doc)
    except RuntimeError:
        print("Невозможно проверить документ.")
        return

    return check.main_check()
