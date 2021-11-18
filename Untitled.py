
import openpyxl
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

pyautogui.press('winleft')
pyautogui.write('chrome')
pyautogui.press('enter')
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

time.sleep(5)

pyautogui.click(x=369, y=304, clicks=2)

pyautogui.click(x=362, y=314)       
pyautogui.click(x=1390, y=196)
pyautogui.click(x=1214, y=592)

time.sleep(5)

tabela = pd.read_excel(r'C:\\Users\hiwan.moreira\Desktop\Vendas - Dez.xlsx')


faturamento = tabela["Valor Final"].sum()
Quantidade = tabela["Quantidade"].sum()
print(tabela)

pyautogui.pause = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

pyautogui.click(x=90, y=203)
pyperclip.copy('hiwan799@gmail.com')
pyautogui.hotkey('ctrl','v')
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v") 
pyautogui.press("tab")

texto = f'''
 O Faturamento do Kemppert foi :R${faturamento:,.2f}
 A Quantidade De igrejas que ele foi :{Quantidade}
 '''
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey('ctrl' , 'enter')









