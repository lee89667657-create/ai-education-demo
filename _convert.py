
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(r"C:\aiproject\학교제안서_AI교육프로그램_예시포함.docx")
doc.SaveAs(r"C:\aiproject\학교제안서_AI교육프로그램_최종.pdf", FileFormat=17)
doc.Close()
word.Quit()
print("PDF saved!")
