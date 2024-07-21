import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

#argument by open browser
chrome_path= "C:/Program Files/Google/Chrome/Application/chrome.exe --start-maximized %s"
#Read data
data = pd.read_excel('files/base_clientes.xlsx')

print(len(data))
for i in range(len(data)):
    print(f"Enviando {i+1}")
    
    cell_phone = '502'+data.loc[i,'Celular'].astype(str)
    message = data.loc[i,'Mensaje']
    #open web browser in chrome
  
    web.get(chrome_path).open(f"https://web.whatsapp.com/send?phone='{cell_phone}'&text='{message}")
    time.sleep(6)
    pg.press('enter')
    time.sleep(2)
    
    #close browser
    pg.hotkey('ctrl','w')
    time.sleep(2)
    
print('fin')