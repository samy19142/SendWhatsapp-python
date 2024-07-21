import pandas as pd
import pyautogui as pg
import webbrowser as web
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox as MessageBox

def send_whatsapp(excel_path):
    chrome_path= "C:/Program Files/Google/Chrome/Application/chrome.exe --start-maximized %s"      #argument by open browser
    data = pd.read_excel(excel_path)   #Read data
    print(len(data))
    for i in range(len(data)):
        print(f"Enviando {i+1}")
        cell_phone = '502'+data.loc[i,'Celular'].astype(str)
        message = data.loc[i,'Mensaje']
        web.get(chrome_path).open(f"https://web.whatsapp.com/send?phone='{cell_phone}'&text='{message}") #open web browser in chrome
        time.sleep(6)
        pg.press('enter')
        time.sleep(2)
        pg.hotkey('ctrl','w') #close browser
        time.sleep(2)
    print('fin')
    
    
def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    archivo_entrada_path = filedialog.askopenfilename(title="Seleccione el archivo de Excel", filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
    if (archivo_entrada_path):
        print(f'el resultado es: {archivo_entrada_path}')
        send_whatsapp(archivo_entrada_path)
        MessageBox.showinfo('Genial!',"El envío finalizó con éxito.")
    else:
        MessageBox.showinfo('Aviso',"No selecciono ningún archivo.")


if __name__ == "__main__":
    seleccionar_archivo()
