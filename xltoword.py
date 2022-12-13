#pyinstaller --onefile xltoword.py cd C:\Users\IMatveev\PycharmProjects\wordchanche\
import os
import zipfile
import re
import numpy as np
import pandas  # +openpyxl
from tkinter import filedialog
import sys
import win32com.client  # pip install pypiwin32
import docx #pip install python-docx
import shutil

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
memcol = 0
isch = 0
usch = 0
found = ""
run_macro(pathword2)
df = pandas.read_excel(pathword2)
doc = docx.Document(pathword)
for para in doc.paragraphs:
    for run in para.runs:
        if len(run.text) > 0:
            if run.text[0] == "{":
                nasr = df[df["metka"] == run.text]["chenge"].to_string(header=False, index=False)
                if nasr == "NaN":
                    print(para.text, "Deleted")
                    delete_paragraph(para)
doc.save(pathwork + "/B.docx")
try:
    os.remove(pathwork + "/B.zip")
except:
    asd = 1
os.rename(pathwork + "/B.docx", pathwork + "/B.zip")
fantasy_zip = zipfile.ZipFile(pathzip)  # extract zip (+need rename docx to zip +need raname vise versa
fantasy_zip.extractall(pathwork + "/B")
fantasy_zip.close()
with open(pathwork + "/B/word/document.xml", 'r', encoding='utf-8') as f:  # save before chenge
    get_all = f.readlines()
print("xml opened")
with open(pathwork + "/B/word/document.xml", 'w', encoding='utf-8') as f:  # look for { and chenge it
    for i in get_all:         # STARTS THE NUMBERING FROM 1 (by default it begins with 0)
        usch = len(i)-1
        for u in i:
            if get_all[isch][usch] == "}":
                mem = 1
                memcol = 0
            if memcol == 1: #замена цвета
                if get_all[isch][usch:usch+7] == "w:fill=":
                    memcol = 0
                    if not get_all[isch][usch+8] == "a":
                        dl = 6
                        #print(get_all[isch][usch - 2:usch + 40])
                        get_all[isch] = get_all[isch][:usch+8] + str(col) + get_all[isch][usch + 8 + dl:]
                    else:
                        dl = 4
                        get_all[isch] = get_all[isch][:usch + 8] + str(col) + get_all[isch][usch + 8 + dl:]
                        #print(get_all[isch][usch-2:usch+40])
                    # print(get_all[isch][usch:usch+50])
            if mem == 1:
                found = get_all[isch][usch] + found
            if get_all[isch][usch] == "{":
                mem = 0
                print(found)
                tx = df[df["metka"] == found]["chenge"].values[0]#header=False,
                try:
                    float(tx)
                    tx = str(tx).replace(".", ",")
                except:
                    asd = 0
                col = df[df["metka"] == found]["color"].values[0]#df[df["metka"] == found]["color"].to_string(header=False, index=False)
                if col == "NaN" or col == "FFFFFF" or col == "":
                    get_all[isch] = get_all[isch][:usch] + tx + get_all[isch][usch + len(found):]
                    found = ""
                else:
                    memcol = 1
                    get_all[isch] = get_all[isch][:usch] + tx + get_all[isch][usch + len(found):]
                    found = ""
            usch = usch - 1
        isch = isch + 1
    f.writelines(get_all)
print("XML chanched")
try:
    os.remove(pathwork + "/B.zip")
except:
    asd = 1
fantasy_zip = zipfile.ZipFile(pathwork + "/B.zip", 'w')
for folder, subfolders, files in os.walk(pathwork + "/B"):
    for file in files:
        fantasy_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), pathwork + "/B"))
fantasy_zip.close()  # transform it to zip
print("zip saved")
name = str(df.iloc[0, 1])
name = "О разв ПМСП в " + name[82:]
try:
    os.remove(pathwork + "/" + name + ".docx")
except:
    asd = 1
os.rename(pathwork + "/B.zip", pathwork + "/" + name + ".docx")
shutil.rmtree(pathwork + "/B/")
print("FINISH")
