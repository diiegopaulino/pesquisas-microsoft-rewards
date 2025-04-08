import pyautogui
import time
from tkinter import messagebox
import keyboard
import threading

# Flag de controle
interromper = False

# Função para monitorar a tecla ESC
def esc_listener():
    global interromper
    keyboard.wait('esc')
    interromper = True

# Mensagem inicial
continuar = messagebox.askyesno(
    "Confirmação",
    "Esta rotina irá abrir o navegador, acessar o site Bing e realizar 38 pesquisas numéricas.\n\n"
    "Para cancelar a qualquer momento, pressione a tecla ESC.\n\nDeseja continuar?"
)

if not continuar:
    messagebox.showinfo("Cancelado", "A rotina foi cancelada pelo usuário.")
    exit()

# Inicia a thread para escutar ESC
esc_thread = threading.Thread(target=esc_listener, daemon=True)
esc_thread.start()

pyautogui.PAUSE = 1

# Abrir o Microsoft Edge
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')
time.sleep(5)

# Entrar no site bing.com
pyautogui.write('bing.com')
pyautogui.press('enter')
time.sleep(5)

# Fazer pesquisas dos números entre 0 e 37
for i in range(38):
    if interromper:
        messagebox.showinfo("Interrompido", "A rotina foi interrompida pelo usuário.")
        exit()
    pyautogui.write(str(i))
    pyautogui.press('enter')
    time.sleep(5)

# Fechar a aba correspondente
pyautogui.hotkey('ctrl', 'w')

# Exibir mensagem de sucesso
messagebox.showinfo("Sucesso", "Pesquisas realizadas com sucesso!")