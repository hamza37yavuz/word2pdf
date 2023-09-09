import os
import win32com.client as win32
from docx2pdf import convert

def convert_doc_to_docx(doc_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    doc.SaveAs2(docx_path, FileFormat=16)  # 16: .docx formatı
    
    doc.Close()
    word.Quit()

# Python dosyasının bulunduğu dizin
current_directory = os.path.dirname(os.path.abspath(__file__))

for filename in os.listdir(current_directory):
    if filename.endswith(".doc"):
        doc_path = os.path.join(current_directory, filename)
        convert_doc_to_docx(doc_path)



# Mevcut dizindeki Word belgelerini PDF'ye dönüştür
current_directory = os.getcwd()  # Mevcut dizini al

convert(current_directory)
