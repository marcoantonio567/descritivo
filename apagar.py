import keyboard
import openpyxl
import pyautogui
import time
import pyperclip

def limpar_area_de_transferencia():
    try:
        pyperclip.copy("")  # Copia uma string vazia para a área de transferência
        print("A área de transferência foi limpa.")
    except Exception as e:
        print(f"Erro ao limpar a área de transferência: {e}")
def verificar_area_de_transferencia():
    try:
        conteudo = pyperclip.paste()
        if conteudo:
            return conteudo
        else:
            print("A área de transferência está vazia.")
    except Exception as e:
        print(f"Erro ao acessar a área de transferência: {e}")
def verificar_hifen(texto):
    if "-" in texto:
        return 'sim'
    else:
        print("O texto não contém um '-' (hífen).")
    
pyautogui.alert("selecione")
time.sleep(1)
for i in range(50):
    pyautogui.hotkey('ctrl','x')
    texto_negativo_positivo = verificar_area_de_transferencia()
    if verificar_hifen(texto_negativo_positivo) == 'sim':
        for i in range(3):
            pyautogui.press('right')
            time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.press("down")
        for i in range(3):
            pyautogui.press('left')
            time.sleep(0.1)
        limpar_area_de_transferencia()
    else:
        for i in range(2):
            pyautogui.press('right')
            time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.press("down")
        for i in range(2):
            pyautogui.press('left')
            time.sleep(0.1)
        limpar_area_de_transferencia()
