from docx import Document

doc = Document('zzz.docx')

tables = doc.tables

print(len(tables))