from cores import *
import pyautogui
import keyboard
import time
from nova_automacao import (quantidade_matriculas , quantidade_propietarios , 
                            data_imoveis , dados_propietarios ,somar_valores,
                            substituir_ponto_por_virgulas, 
                            layout_geral_mapa,apagar_pycache,
                            valores_mercado_automaticos,valores_liquidacao_automaticos)


def time_next():
    time.sleep(0.3)
def tempo():
    time.sleep(0.5)
def clicar_no_centro():
    # Obtem a largura e altura da tela
    largura, altura = pyautogui.size()

    # Calcula as coordenadas do centro
    centro_x = largura // 2
    centro_y = altura // 2

    # Aguarda 2 segundos antes de clicar para o usuário se preparar
    print(f"{DARK_GREEN}Clicando no centro da tela em 2 segundos...{RESET}")
    
    
    # Move o mouse para o centro da tela e clica
    pyautogui.moveTo(centro_x, centro_y,duration=0.5)
    pyautogui.click()
def seta_baixo():
    pyautogui.press("down")
def tab():
    pyautogui.press("tab")
    time_next()
def proxima_linha():
    pyautogui.hotkey('alt','6')
def desfazer_ultima_linha():
    pyautogui.hotkey('ctrl','z')

pyautogui.alert("habilite o conteudo e aperte ok")
clicar_no_centro()
time.sleep(1)
pyautogui.hotkey('ctrl','home')
tempo()
if quantidade_propietarios >1:
    for i in range(4):
        tab()#aqui ele ta indo pro prencher o primeiro propietario
    tempo()
    for dado in dados_propietarios:
        if len(dado[1]) == 14:
            verificacao_cpf_cnpj = "CPF: "
        else:
            verificacao_cpf_cnpj = "CNPJ: "
        keyboard.write('Proprietário:',delay=0.002)
        tab()
        keyboard.write(str(dado[0]),delay=0.002)#escrever o nome do propietario
        tab()
        keyboard.write(verificacao_cpf_cnpj+str(dado[1]),delay=0.002)
        proxima_linha()
        tempo()
    desfazer_ultima_linha()
    tempo()
    for i in range(15):#indo ate aas matriculas
        tab()
else:   
    for i in range(21):#indo ate aas matriculas
        tab()
if quantidade_matriculas>1:
    
    tempo()
    for fazenda in data_imoveis:
        new_Area_ha = substituir_ponto_por_virgulas(fazenda[2],casas_decimais=4)
        new_Area_construida = substituir_ponto_por_virgulas(fazenda[3],casas_decimais=2)
        keyboard.write(str(fazenda[0]),delay=0.002)
        tab()
        keyboard.write(str(fazenda[1]),delay=0.002)
        tab()
        keyboard.write(new_Area_ha,delay=0.002)
        tab()
        keyboard.write(new_Area_construida,delay=0.002)
        proxima_linha()
        tempo()
    desfazer_ultima_linha()
    tempo()
    #aqui to indo ate os valores de mercado
    pyautogui.alert("selecione os valores de mercado e parte ok")
    for i in range(4):#aqui to indo pra definitivamente escrever
        tab()
    tempo()
    for fazenda in data_imoveis:
        valor_mercado_formato_escrever = fazenda[4]
        valor_liquidacao_formatado_escrever = fazenda[5]
        keyboard.write(str(fazenda[0]),delay=0.002)
        tab()
        keyboard.write(valor_mercado_formato_escrever,delay=0.002)
        tab()
        keyboard.write(valor_liquidacao_formatado_escrever,delay=0.002)
        proxima_linha()
        tempo()
    soma_valor_mercado = somar_valores(valores_mercado_automaticos)
    soma_valor_liquidacao = somar_valores(valores_liquidacao_automaticos)
    keyboard.write('TOTAL',delay=0.002)
    tab()
    keyboard.write(soma_valor_mercado,delay=0.002)
    tab()
    keyboard.write(soma_valor_liquidacao,delay=0.002)
    tempo()
    

pyautogui.alert("selecione a capa por favor")
tempo()
pyautogui.hotkey("alt","7")#aqui é pra ele alterar a imagem
time.sleep(2)#esperar o arquivo carregar
"""for i in range(1):
    tab()#aqui ele vai ir ate em procurar"""
pyautogui.press('enter')
time.sleep(1)
keyboard.write(str(layout_geral_mapa),delay=0.001)
time.sleep(1)
pyautogui.press('enter')
tempo()
apagar_pycache()
pyautogui.alert("terminooou! ⊂(◉‿◉)つ\nproponha oque voce acha que pode ser menlhorado ")



