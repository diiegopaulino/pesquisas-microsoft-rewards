import pyautogui
import time
import tkinter as tk
from tkinter import messagebox

import win32com.client as win32

pyautogui.PAUSE = 1

# Abrir o Microsoft Edge
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')

# Esperar o navegador abrir
time.sleep(5)

# Entrar no site bing.com
pyautogui.write('bing.com')
pyautogui.press('enter')

# Esperar o site carregar
time.sleep(5)

# Fazer pesquisas dos números entre 0 e 37
for i in range(38):
    pyautogui.write(str(i))
    pyautogui.press('enter')
    time.sleep(5)

# Fechar a aba correspondente
pyautogui.hotkey('ctrl', 'w')

# Exibir mensagem de sucesso
messagebox.showinfo("Sucesso", "Pesquisas realizadas com sucesso!")