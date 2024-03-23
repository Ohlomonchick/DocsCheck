import aspose.words as aw

lic = aw.License()
try:
    lic.set_license("Aspose.WordsforPythonvia.NET.lic")
    print("License set successfully.")
except RuntimeError as err:
    # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
    print("\nThere was an error setting the license:", err)
doc = aw.Document("samples/Система_для_автоматического_перелистывания_нот_на_планшете_актуальное.docx")


from checker import BaseChecker

checker = BaseChecker(doc)



verdict = checker.main_check()
for message in verdict.messages:
    print(message)