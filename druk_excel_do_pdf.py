import os
import shutil
from win32com import client
from PyPDF2 import PdfMerger

'''
Program takes excel files from one dir and
creates pdfs from them in another dir, then
merge the pdf files into one
'''

# check if pliki_pdf folder exists, if true delete it
if os.path.exists('D:/pyproject/Drukarka_PDF/pliki_pdf'):
    shutil.rmtree('D:/pyproject/Drukarka_PDF/pliki_pdf')

# create pliki_pdf directory
os.mkdir('D:/pyproject/Drukarka_PDF/pliki_pdf')

# saving name of every file in pliki_excela directory to a list
list_of_files = os.listdir('D:/pyproject/Drukarka_PDF/pliki_excela')
print(list_of_files)

# Open Microsoft Excel
excel = client.Dispatch('Excel.Application')

# open every file in directory and create pdf from it
for i in list_of_files:

    sheets = excel.Workbooks.Open(f'D:/pyproject/Drukarka_PDF/pliki_excela/{i}')
    work_sheets = sheets.Worksheets[0]

    # convert to PDF File
    work_sheets.ExportAsFixedFormat(0, f'D:/pyproject/Drukarka_PDF/pliki_pdf/{i}')

    # close excel file
    sheets.Close()

# close Microsoft Excel
excel.Quit()

# creating list of files in pliki_pdf directory
list_of_pdf_files = os.listdir('D:/pyproject/Drukarka_PDF/pliki_pdf')
print(list_of_pdf_files)

# merging files
merger = PdfMerger()

for i in list_of_pdf_files:
    merger.append(f'D:/pyproject/Drukarka_PDF/pliki_pdf/{i}')

merger.write('scalone.pdf')
merger.close()
