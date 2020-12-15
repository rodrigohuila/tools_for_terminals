#! python3
#! /usr/bin/python3

import os
import PyPDF2
from tkinter import filedialog as FileDialog
from tkinter import messagebox as MessageBox
from os import system
from os.path import basename

path = r"C:\Users\hector.huila\Desktop\New VGM"
path2 = r"C:\Users\hector.huila\Desktop\New VGM"
contentAll = ''


def clear():  # BORRAR PANTALLA
    if os.name == "nt":
        os.system("cls")
    else:
        os.system("clear")


def openFile():  # DIALOG TO CHOOSE A FILE
    pdfName = FileDialog.askopenfilename(
        initialdir=path,
        filetypes=(('pdf', '*.pdf'), ('all files', '*.*')),
        title='Escoja el archivo de imagen a dividir y renombrar'
    )
    name = basename(pdfName)
    print('\nEl nombre del archivo inicial es: ' + name + '\n')
    return(pdfName)



def extract_text(pdfReader, pageNum): # EXTRACT TEXT FROM PDF FILE
    pageObj = pdfReader.getPage(pageNum)  # Read the content of the first page
    cadena = pageObj.extractText()
    print ('\nEl contenido de la página No: ' + str(pageNum) + ' del PDF es el siguiente: \n' + cadena)
    #return(cadena)
    text = read_texto(cadena)
    return(text)


def read_texto(cadena):  # READ THE ESPECIFIC SRRING THAT WE NEED
    startWord = cadena.index('TERMINAL')  # Ingresamos el texto inicial de la búsqueda
    endWord = cadena.index('%')# Ingresamos el texto final de la búsqueda

    subcadena = cadena[startWord:endWord].replace('\n', ' ').replace('TERMINAL', '')
    #newString = subcadena.replace('\n', ' ')
    subcadena = subcadena.split(' ')  # Split string to capitalize word by word
        
    finalName = []
    for word in subcadena:
        if len(word) > 2 and (word.isalpha() == True):
            finalName.append(word.lower())
        elif len(word) <= 2 and (word.isalpha() == True):
            finalName.append(word.lower())
        elif len(word) >= 2 and (word.isdigit() == True):
            word2 = 'Peso ' + word
            finalName.append(word2)

    text = '_'.join(finalName)
    return(subcadena)
    


def main():
    clear()
    pdfName = openFile() # Open the pdf file
    pdfFileObj = open( pdfName, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) # Read info
    content = extract_text(pdfReader, 0)
    print ('\nFueron Procesadas: ' + str(pdfReader.numPages) + ' páginas')
    print (content)
    
        
if __name__ == '__main__':
    main()    
    
os.system('pause') # Press a key to continue        
