from docsCheck.runners import run_check


verdict = run_check(
    "samples/Система_для_автоматического_перелистывания_нот_на_планшете_актуальное.docx",
    doc_type="ТЗ"
)
for message in verdict.messages:
    print(message)