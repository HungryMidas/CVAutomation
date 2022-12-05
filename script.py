import pyautogui
import pandas
import datetime
import time
from docx2pdf import convert
from docx import Document
import os 


# Author @inforkgodara

# Read data from excel
# "C:\Users\minhh\OneDrive\Documents\Professional\CV Automation\generate-bulk-documents-python\data.xlsx"
# excel_ref = "C:\Users\minhh\OneDrive\Documents\Professional\CV Automation\generate-bulk-documents-python\data.xlsx"
excel_data = pandas.read_excel('data.xlsx', sheet_name='Recipient Details')
count = 0
directory = 'generated'

def replaceWord(oldString, newString, paragraph):
    if oldString in paragraph:
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if oldString in inline[i].text:
                text = inline[i].text.replace(oldString, newString)
                inline[i].text = text

# "C:\Users\minhh\OneDrive\Documents\Professional\CV Automation\generate-bulk-documents-python\template.docx"
# template_ref = r'C:\Users\minhh\OneDrive\Documents\Professional\CV Automation\generate-bulk-documents-python\template.docx'
# Iterate excel rows till to finish
today = str(datetime.datetime.now().strftime("%m/%d/%Y"))
# print(today)
for column in excel_data['Recipient'].tolist():
    # file = open(template_ref, 'rb')
    document = Document('template.docx')
    doc = document
    companyName = excel_data['Company'][count]
    empName = excel_data['Recipient'][count]
    for p in doc.paragraphs:
        replaceWord('REASON', excel_data['Reason'][count], p.text)
        replaceWord('DATE', today, p.text)
        replaceWord('POSITION', excel_data['Position'][count], p.text)
        replaceWord('HIRING MANAGER', excel_data['Recipient'][count], p.text)
        replaceWord('HIRING MANAGER FIRST NAME', excel_data['Recipient First Name'][count], p.text)
        replaceWord('TITLE', excel_data['Recipient Title'][count], p.text)
        replaceWord('COMPANY', excel_data['Company'][count], p.text)
        replaceWord('STREET ADDRESS', excel_data['Recipient Street Address'][count], p.text)
        replaceWord('CITY, ST ZIP CODE', str(excel_data['Recipient City, ST ZIP Code'][count]), p.text)

    try:
        path = os.getcwd()+"/"+directory+"/"+companyName
        os.mkdir(path)
    except OSError:
        a = 10

    doc.save(os.getcwd()+"/"+directory+"/"+companyName+"/"+companyName+' Letter.docx')
    convert("generated/" + companyName)
    print("Letter generated for " + companyName)
    count = count + 1

print("Total letters are created " + str(count))