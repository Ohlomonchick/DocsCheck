import aspose.words as aw


lic = aw.License()
try:
    lic.set_license("Aspose.WordsforPythonvia.NET.lic")
    print("License set successfully.")
except RuntimeError as err:
    # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
    print("\nThere was an error setting the license:", err)
doc = aw.Document("samples/Система_для_автоматического_перелистывания_нот_на_планшете_актуальное.docx")

# for field in doc.range.fields:
#
#     if (field.type == aw.fields.FieldType.FIELD_HYPERLINK):
#
#         hyperlink = field.as_field_hyperlink()
#         if (hyperlink.sub_address != None and hyperlink.sub_address.find("_Toc") == 0):
#             tocItem = field.start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
#
#             print(tocItem.to_string(aw.SaveFormat.TEXT).strip())
#             print("------------------")
#
#             bm = doc.range.bookmarks.get_by_name(hyperlink.sub_address)
#             try:
#                 pointer = bm.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
#
#                 print(pointer.to_string(aw.SaveFormat.TEXT))
#             except Exception:
#                 print("cant read ----")

# Document doc = new Document(“C:\Temp\in.doc”);
#
# LayoutCollector layoutCollector = new LayoutCollector(doc);
#
# doc.updatePageLayout();
#

def check_if_on_page(page_number, text):
    extractedPage = doc.extract_pages(page_number, 1)
    print(extractedPage.to_string(aw.SaveFormat.TEXT))

from checker import BaseChecker

checker = BaseChecker(doc)
print(checker.check_certification_page().messages)
