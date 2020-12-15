#! python3

import os
from tkinter import filedialog as FileDialog
from tkinter import messagebox as MessageBox
from os.path import basename
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy

pathFile = r"S:/OPS/Marítimas/CARPETAS MOTONAVES"


def clear():  # ERASE SCREEN
    if os.name == "nt":      # Windows
        os.system("cls")
    else:
        os.system("clear")   # Linux


def openFile():  # DIALOG TO CHOOSE A FILE
    sealFile = FileDialog.askopenfilename(
        initialdir=pathFile,
        filetypes=(('xls', '*.xls'), ('all files', '*.*')),
        title='Escoja el archivo de Tarja a procesar'
    )
    print('\nEl nombre del archivo inicial es: ' + basename(sealFile) + '\n')
    #return(basename(sealFile))
    return(sealFile)


def main():  # MAIN
    clear()
    sealFile = openFile()  # Name of the file to choose

    book = xlrd.open_workbook(sealFile)
    sheet = book.sheet_by_index(1)
    book = copy(book)
    sheet2 = book.get_sheet(1)  # same sheet in order to can save

    # Loop through the sheet
    for i in range(sheet.nrows):
        sello1 = (repr(sheet.cell_value(i, 7))[1:7])
        sello2 = (repr(sheet.cell_value(i, 8)))

        if (sello2[1:3]) == 'ML':
            if sello2.find(sello1) > 0:
                # background color
                style = xlwt.easyxf(
                    'pattern: pattern solid, fore_colour yellow')
                sheet2.write(i, 7, '0', style)  # Modify cell
                # sheet2.write(i, 8, (sello2.replace("'", "")),
                #             style)  # Modify cell

    book.save(sealFile)  # Save


if __name__ == '__main__':
    main()
