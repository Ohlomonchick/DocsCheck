from src.docsCheck.runners import run_check
from src.docsCheck.__main__ import print_verdict

# "C:\\Users\\Dmitry\\OneDrive\\ПЗ Курсач Проскурин.docx",
verdict = run_check(
    "samples/Система_для_автоматического_перелистывания_нот_на_планшете_актуальное.docx",
    doc_type="ТП"
)
print_verdict(verdict)