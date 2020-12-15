#! python3

from os.path import basename
import pyautogui
import pyperclip
import re
import os
import time
from datetime import datetime, date, timedelta


now = datetime.now()
dateformat = now.strftime('%d-%m-%Y_%H-%M')
getpath = os.getcwd() # directory where the script is executed
pathfile = r'"{}\Existencia_{}.xls"'.format(getpath, dateformat)


def clear():  # ERASE SCREEN
    if os.name == 'nt':      # Windows
        os.system('cls')
    else:
        os.system('clear')   # Linux


def mouse_selection(opc):
    # pyautogui.displayMousePosition() # Know where the mouse is in real time
    while True:
        try:
            if opc == 1:
                # Monitor 1
                pyautogui.click(488, 34) # Conusltas monitor 2
                pyautogui.click(534, 77) # Contenedores en terminal monitor 2
                time.sleep(2.5)
                pyautogui.click(747, 203) # Exportar monitor 2
                pyautogui.click(835, 532) # Posición para copiar ruta monitor 2
                pyautogui.typewrite(pathfile, interval=0.05) # Write the path monitor 2
                pyautogui.click(911, 570) # Aceptar monitor 2
                time.sleep(3.5)
                pyautogui.click(1005, 568) # Cancelar monitor 2
                pyautogui.click(1568, 167) # Cerrar monitor 2
                print (f'Archivo Existencia_{dateformat} ha sido guardado exitosamente en {getpath}')
                break

            elif opc ==2:
                # Monitor 2
                pyautogui.click(2423, 35) # Conusltas monitor 2
                pyautogui.click(2442, 81) # Contenedores en terminal monitor 2
                time.sleep(2.5)
                pyautogui.click(2661, 201) # Exportar monitor 2
                pyautogui.click(2755, 528) # Posición para copiar ruta monitor 2
                pyautogui.typewrite(pathfile, interval=0.05) # Write the path monitor 2
                pyautogui.click(2834, 568) # Aceptar monitor 2
                time.sleep(3.5)
                pyautogui.click(2911, 570) # Cancelar monitor 2
                pyautogui.click(3486, 163) # Cerrar monitor 2
                print (f'Archivo Existencia_{dateformat} ha sido guardado exitosamente en {getpath}')
                break

        except ValueError:
            ornament()
            print ('Opción invalida')
            clear()
            main()
        

def ornament():
    print("_" * 50)


def menu1():
    ornament()
    print("Yard Managment")
    ornament()
      
    

def main():
    menu1()
    print ()
    opc = input ('Monitor1 or Monitor2: ')
    mouse_selection(int(opc))



if __name__ == '__main__':
   main()


os.system("pause") # Press a key to continue 
