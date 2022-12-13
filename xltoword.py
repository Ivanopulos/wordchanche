import os
import zipfile
import pandas  # +openpyxl
from tkinter import filedialog
import sys
import win32com.client  # pip install pypiwin32
import docx #pip install python-docx

print(os.environ.get( "USERNAME" ))
def run_macro(name):
    print('macro')
    if os.path.exists(name):
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Workbooks.Open(Filename=name, ReadOnly=1)
        xl.Application.Run("'C:\\Users\\" + os.environ.get( "USERNAME" ) + "\\AppData\\Roaming\\Microsoft\\AddIns\\Ivax.xlam'!цвета")
        xl.Application.Quit() # Comment this out if your excel script closes
        del xl
    print('File refreshed!')
def resource_path(relative):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(relative)
def delete_paragraph(paragraph):   # Delete stroke+
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
print("ворд")
pathword = filedialog.askopenfilename()
pathwork = os.path.dirname(pathword)
pathzip = pathwork + "/B.zip"
pathword2 = pathwork + "/Шаблон написания справки.xlsx"
mem = 0
isch = 0
usch = 0
found = ""
run_macro(pathword2)
df = pandas.read_excel(pathword2)

doc = docx.Document(pathword)
outt=0
for para in doc.paragraphs:
    for run in para.runs:
        print(run.text)
        #if run.text == ""//////////////old_info:
        #        outt = 1
        #        if new_info == "":
        #            delete_paragraph(para)
        #   run.text = run.text.replace(old_info, new_info)  # информация о замене

