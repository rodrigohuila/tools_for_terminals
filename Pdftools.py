#! python3

try:
    from PIL import Image
except ImportError:
    import Image
from pdf2image import convert_from_path
import img2pdf
from PIL import Image
import pytesseract
import os
import shutil
import zipfile
from tkinter import filedialog as FileDialog
from tkinter import messagebox as MessageBox
from os import system
from os.path import basename
import re
import PyPDF2
#ATTACHMENT
import smtplib, sys, openpyxl, email
from email import encoders
from string import Template
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# If you need to assign tesseract to path
# pytesseract.pytesseract.tesseract_cmd = 'C:\Users\victo\AppData\Local\Tesseract-OCR\tesseract.exe'
path = r"D:/Users/victo/Downloads/"
pathFile = r"D:/Users/victo/Desktop/RenamePDF/"
path2 = r"D:/Users/victo/Desktop/RenamePDF/"
ext = 'jpg'
zipName = 'Diplomas.zip'
#ATTACHMENT
MY_ADDRESS = 'victoria.soto@co.g4s.com'
pathCartas = r"D:/Users/victo/Desktop/Carta_Examen_Simetric/"
ext2 = '.pdf'
contacts = 'Cartas SIMETRIC.xlsx'
msgTemplate = r"D:/Users/victo/Desktop/RenamePDF/Scripts/message.txt" #'message.txt'


def clear():  # BORRAR PANTALLA
    if os.name == "nt":
        os.system("cls")
    else:
        os.system("clear")

def openFile():  # DIALOG TO CHOOSE A FILE
    pdfName = FileDialog.askopenfilename(
        initialdir=path,
        filetypes=(('Archivos PDF', '*.pdf'), ('all files', '*.*')),
        title='Escoja el archivo de imagenPDF a dividir y renombrar'
    )
    name = basename(pdfName)
    print('\nEl nombre del archivo inicial es: ' + name + '\n')
    return(pdfName)

def getDirectory():  # DIALOG TO CHOOSE A DIRECTORY
    newDirectory = FileDialog.askdirectory(initialdir=pathFile, title='Escoja el directorio o Carpeta en donde se encuentran las CARTAS SIMETRIC')
    directoryPath = newDirectory
    return(directoryPath)           

def convert_image(imageName, pathFile, ext):  # CONVERT IMAGEN A SELECT EXT
    lstFiles = []

    convert_from_path(imageName, output_folder=pathFile,
                      fmt=ext, output_file='Diploma')
    
    for fichero in os.listdir(pathFile):
        (nombreFichero, extension) = os.path.splitext(fichero)
        if(extension == "." + ext):
            lstFiles.append(nombreFichero+extension)
    return(lstFiles)

def ocr_core(fileName):  # OCR PROCESSING OF IMAGES
    text = pytesseract.image_to_string(Image.open(
        fileName))  # Pillow's Image class open image and pytesseract detect strings
    return text

def read_image(cadena):  # READ LETTERS AND STRINGS IN THE IMAGE WITH GOOGLE OCR
    startWord = cadena.index(
        'CERTIFICA')  # Ingresamos el texto inicial de la búsqueda
    # Ingresamos el texto final de la búsqueda
    endWord = cadena.index('CON UNA')

    subcadena = cadena[startWord:endWord]
    newString = subcadena.lstrip("CERTIFICA QUE").lstrip("CERTIFICA").lstrip(
        "QUE").lstrip(':').replace('\n', ' ').replace('.', '').replace('ASISTIO Y APROBO EL CURSO DE', 'Curso de').strip().rstrip("CON UN").rstrip().rstrip().replace(' ', '_').replace(':', '').lower()
     
    newString = newString.split('_')  # Split string to capitalize word by word
        
    newString2 = []
    for word in newString:
        if len(word) > 2 and (word.isalpha() == True) or (word.isdigit() == True):
            newString2.append(word.capitalize())
        elif len(word) <= 2 and (word.isalpha() == True) or (word.isdigit() == True):
            newString2.append(word)
    newString2 = '_'.join(newString2)
    return(newString2)

def rename_file(name, finalName):  # RENAME FILES
    archivo = name
    nombreNuevo = (finalName + "." + ext)
    os.rename(archivo, nombreNuevo)
    return(nombreNuevo)

def convert_pdf(nameFile, finalName):  # CONVERT FILE FROM A SELECTED EXT TO PDF
    image = Image.open(nameFile)  # opening image
    # converting into chunks using img2pdf
    pdf_bytes = img2pdf.convert(image.filename)
    file = open(finalName + ".pdf", "wb")  # opening or creating pdf file
    file.write(pdf_bytes)  # writing pdf files with chunks
    image.close()  # closing image file
    file.close()  # closing pdf file
    print("Successfully made pdf file: " + file.name)  # output

def compress_file(newExt):  # COMPRESS FILES IN A ZIP FILE
    fileszip = zipfile.ZipFile(pathFile + zipName, 'w')

    for file in os.listdir(pathFile):
        if file.endswith(newExt):
            fileszip.write(file, compress_type=zipfile.ZIP_DEFLATED)
            os.unlink(file)  # Erase the file using shutil module
    fileszip.close()


## CARTAS SIMETRIC
def pdf_splitter(pdfName, pdfReader, pathFile): #SPLIT PDF FILE
    for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter = PyPDF2.PdfFileWriter()
        pdfWriter.addPage(pageObj)
        finalName = extract_text(pdfReader, pageNum)  #Enviar pdf para extraer texto
        output_filename = '{}.pdf'.format(
            finalName)
        os.chdir(pathFile)
        with open(output_filename, 'wb') as out:
            pdfWriter.write(out)
        print('Created: {}'.format(output_filename))
        
def extract_text(pdfReader, pageNum): # EXTRACT TEXT FROM PDF FILE !!!! Devuelve a la anteriror función
    pageObj = pdfReader.getPage(pageNum)  # Read the content of the first page
    cadena = pageObj.extractText()
    #print ('\nEl contenido de la página No: ' + str(pageNum) + 'del PDF es el siguiente: \n' + content)
    #return(content)
    finalName = read_texto(cadena)
    return(finalName)

def read_texto(cadena):  # READ THE ESPECIFIC SRRING THAT WE NEED !!!!! Devuelve a la anteriror función
    startWord = cadena.index(
        'Señor (a)')  # Ingresamos el texto inicial de la búsqueda
    # Ingresamos el texto final de la búsqueda
    endWord = cadena.index('La Directora ')

    subcadena = cadena[startWord:endWord].replace('\n', ' ').replace('Señor', '')
    #newString = subcadena.replace('\n', ' ')
    subcadena = subcadena.split(' ')  # Split string to capitalize word by word
        
    finalName = []
    for word in subcadena:
        if len(word) > 2 and (word.isalpha() == True):
            finalName.append(word.capitalize())
        elif len(word) <= 2 and (word.isalpha() == True):
            finalName.append(word.lower())
        elif len(word) > 2 and (word.isdigit() == True):
            word2 = 'CC_' + word
            finalName.append(word2)

    finalName = '_'.join(finalName)
    return(finalName)

def pdf_splitter2(pdfName, pdfReader): # ONLY SPLIT A PDF DOCUMENT
    os.chdir(path2)
    pathFile = newCarpeta(pdfName, path2) # Nueva Carpeta con el nombre del archivo a procesar
    newDir = (os.path.splitext(basename(pathFile))[0])

    for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter = PyPDF2.PdfFileWriter()
        pdfWriter.addPage(pageObj)
        output_filename = 'page_{}_{}'.format(
            pageNum+1, basename(pdfName))
        os.chdir(path2 + newDir) # Guardar en una carpeta con el mismo nombre del archivo requerido
        with open(output_filename, 'wb') as out:
            pdfWriter.write(out)
        print('Created: {}'.format(output_filename))            
    return(pageObj)    

def newCarpeta(pdfName, path2): # Nueva Carpeta con el nombre del archivo a procesar
    os.chdir(path2)
    if os.path.isdir((os.path.splitext(basename(pdfName))[0])):
        MessageBox.showwarning(
            "Advertencia","""Hay una carpeta con el mismo nombre del PDF, es decir, ya se había dividido este PDF o un archivo con el mismo nombre.\n
            Los archivos se sobreescribiran""")
        newDir = (os.path.splitext(basename(pdfName))[0])
        pathFile = path2 + newDir                      
    else:    
        os.mkdir(os.path.splitext(basename(pdfName))[0])
        newDir = (os.path.splitext(basename(pdfName))[0]) # Crear una carpeta con el mismo nombre del archivo requerido
        pathFile = path2 + newDir
    return(pathFile)    


# GET THE ALL THE NAME OF THE EXCEL LIST
def get_contacts(filename): 
    """
    Return two lists names, emails containing names and email addresses
    read from a file specified by filename.
    """
    names = []
    emails = []
    wb = openpyxl.load_workbook(pathCartas + filename)
    sheet = wb.worksheets[0]
    print('_'*50)
    print('Escoja las Cartas Simetric que desea enviar')
    print('_'*50)
    print()
    desde = int(input('Digite el Item de la casilla de excel desde: '))
    hasta = int(input('Digite hasta: '))
    
    for r in range(desde + 1, hasta + 2): # + 1 PARA QUE TOME EL VALOR DESEADO O SI NO TOMA MENOS 1
          apellido=(sheet.cell(row=r, column=4).value)
          name=(sheet.cell(row=r, column=5).value)
          names.append(apellido + ' ' + name)
          email=(sheet.cell(row=r, column=11).value)
          emails.append(email)        
        
    return names, emails


def read_template(filename): 
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def search_files(names, path):
    """
    Return a list of names ONLY WITH files to attach to the msg
    """
    os.chdir(path)
    files = []
    newNames = []
    #Search the names against the files in the directory
    #for i in range (0, len(names)):                
    for filename in os.listdir(path):        
        for i in range (0, len(names)):  
            name = names[i].lower()                          
            if filename.endswith(ext2):
                fileNewName = str(filename.replace('_', ' ').lower()) #Convert to string an lowercase
                if fileNewName.find(name) >=0:
                    files.append(filename)
                    newNames.append(names[i])
                    i += 1
                #elif name.find(fileNewName) >=0:
                    #files.append(filename)
                    #check = True
                
    set_difference = set(names) - set(newNames) # Compre the two list to find names without attach or with littles problems
    list_difference = list(set_difference)
    for i in range (0, len(list_difference)):
        print('*'*100)
        print('La persona {} no tiene archivo disponible revise y/o envíe manualmente' .format(list_difference[i])) 
        #print(list_difference) # Print the differences             
                    
    files.sort()
    newNames.sort()
    return newNames, files


def search_emails(newName, name, emails):
    """
    Return a FINAL list of names and FINAL emails only if they have a attach, 
    """
    newEmails = []
    newNames = []
    for i in range (0, len(newName)):
        for j in range (0, len(name)):         
            if newName[i] == name[j]:
                newNames.append(newName[i]) 
                newEmails.append(emails[j])            
    return newNames, newEmails


def imprimir():
    print("___________________________________")
    print("Programa para manejo de PDF")
    print("___________________________________")
    print()
    print("Opciones disponibles:")
    print("1. Diplomas")
    print("2. Cartas Simetric")
    print("3. Enviar Cartas Simetric via GMAIL")
    print("4. Dividr un PDF")
    print("5. Unir Imagenes JPEG en un PDF")
    print("0. Salir")
    print()



def main(): # MAIN
    clear()
    while True:
        imprimir()
        try: 
            option = int(input("Seleccione una opcion: "))

            if option == 1: # DIPLOMAS
                pdfName = openFile()
                pathFile = newCarpeta(pdfName, path2) # Nueva Carpeta con el nombre del archivo a procesar
                
                fileList = convert_image(pdfName, pathFile, ext)
                os.chdir(pathFile)  # Go to the Directory
                for i in range(0, len(fileList)):
                    allTextfromImage = ocr_core(fileList[i])  # All text
                    subText = read_image(allTextfromImage)  # Only text with need
                    # Rename image file with text
                    imageFile = rename_file(fileList[i], subText)
                    convert_pdf(imageFile, subText)  # convert to PDF again
                    os.unlink(imageFile)  # Erase the images files using shutil module

                print('\nTotal de archivo renombrados y comprimidos: ' + str(len(fileList)))
                #MessageBox.showinfo('Total de archivo renombrados y comprimidos:', str(len(fileList)) )
                #compress_file('pdf')  # compress all pdf files
                #break
            elif option == 2: # CARTAS
                pdfName = openFile() # Open the pdf file
                pdfFileObj = open( pdfName, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj) # Read info
                pathFile = newCarpeta(pdfName, path2) # Nueva Carpeta con el nombre del archivo a procesar
                pdf_splitter(pdfName, pdfReader, pathFile)
                print ('\nFueron Procesadas: ' + str(pdfReader.numPages) + ' páginas')
                #break
            elif option == 3: # ENVIAR CARTAS POR GMAIL
                clear()
                PASSWORD = ('xqwdzjbamvfiztnt') #input('Digite la clave de aplicaciones para Gmail: ')                            
                os.chdir(pathCartas)
                initialNames, initialEmails = get_contacts(contacts)  # read the initial contacts
                message_template = read_template(msgTemplate) # Cuerpo del mensaje
                pathdirectory = getDirectory() # get Directory where the Cartas are
                names, attach = search_files(initialNames, pathdirectory) # Math only contact that have attach file
                finalNames, mails = search_emails(names, initialNames, initialEmails) #get email ordered like the attached
                
                s = smtplib.SMTP(host='smtp.gmail.com', port=587) # set up the SMTP server
                s.starttls()
                s.login(MY_ADDRESS, PASSWORD)
                
                i = 0 #Iniciar la varaible i en ceros
                for finalName, mail in zip(finalNames, mails): # FOR EACH CONTACT, SEND THE EMAIL:
                    msg = MIMEMultipart()       # create a message                            
                    message = message_template.substitute(PERSON_NAME=finalName.title()) # Match the final names to the message template
                           
                    # setup the parameters of the message
                    msg['From'] = MY_ADDRESS
                    msg['To'] = mail
                    msg['Subject'] = "Exámenes psicofísicos (SIMETRIC)"

                    msg.attach(MIMEText(message, 'plain')) # add in the message body

                    #ATTACH FILES TO THE EMAIL
                    filenameAtach = attach[i]
                    if filenameAtach != None:
                        with open(filenameAtach, "rb") as attachment: # Open PDF file in binary mode
                            # Add file as application/octet-stream
                            # Email client can usually download this automatically as attachment
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(attachment.read())
                        encoders.encode_base64(part) # Encode file in ASCII characters to send by email    
                        part.add_header('Content-Disposition', 'attachment', filename=filenameAtach) # Add header as key/value pair to attachment part
                        msg.attach(part) # Add attachment to message and convert message to string
                        sendmailStatus = s.send_message(msg) # SEND THE MESSAGE WITH ATTACHED.
                        del msg
                           
                        #Status message sent
                        if sendmailStatus != {}:
                            print ('\nThere was a problem sending the email to {}: {}' .format(mail, s.send_message))
                        else:            
                            print ('\nThe email to {} was sent correctly to:\n{}\nwith the attached:\n{}' .format(mail, finalName, filenameAtach))
                        i += 1
                    else: 
                        # SEND THE MESSAGE WITHOUT ATTACHED.
                        #s.send_message(msg)
                        del msg
                        #Do not send a message
                s.quit()  # Terminate the SMTP session and close the connection
                #break
            elif option == 4: # SLPIT PDF
                pdfName = openFile()
                pdfFileObj = open( pdfName, 'rb')  # Open the pdf file
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj) # Read info
                pdf_splitter2(pdfName, pdfReader)
                print ('\nEl documento tiene: ' + str(pdfReader.numPages) + ' páginas')
                #break
            elif option == 5: # UNIR IMAGENES EN UN PDF
                clear()
                # convert all files ending in .jpg inside a directory
                dirname = getDirectory()
                os.chdir(dirname)
                with open((str(input("Digite el nombre del PDF final:\n")) + ".pdf"), "wb") as f:
                    imgs = []
                    for fname in os.listdir(dirname):
                        if not fname.endswith(".jpeg"):
                            continue
                        path = os.path.join(dirname, fname)
                        if os.path.isdir(path):
                            continue
                        imgs.append(path)
                    f.write(img2pdf.convert(imgs))
                #break    
            elif option == 0: # SALIR
                break    
            else:
                print()
                clear()
                print('Error, solo de aceptan numeros del 0 al 4')
                
        except ValueError:
                print("Error, ingrese solamente numeros")            


    
if __name__ == '__main__':
    main()

#os.system('pause') # Press a key to continue
