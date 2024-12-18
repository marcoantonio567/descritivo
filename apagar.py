import tkinter as tk
from tkinter import filedialog


def selecionar_imagens_dos_maps():
    # Cria uma janela oculta do Tkinter
    root = tk.Tk()
    root.withdraw()

    # Abre o seletor de arquivos e permite escolher múltiplas imagens
    caminhos = filedialog.askopenfilenames(
        title="Selecione os MAPAS",
        filetypes=[("Imagens PNG", "*.png")]
    )

    # Retorna os caminhos selecionados como uma lista
    return list(caminhos)


def encontrar_nomes(lista, nomes):
 
    resultados = {}
    for nome in nomes:
        resultado = next((item for item in lista if nome.lower() in str(item).lower()), None)
        resultados[nome] = resultado
    return resultados

# Exemplo de uso:
lista_de_imagens = selecionar_imagens_dos_maps()

nomes_procurados = ["layout geral", "PEDOLOGIA", "vegetação", "bacia" ,"declividade"]
resultados = encontrar_nomes(lista_de_imagens,nomes_procurados)
for nome, item_encontrado in resultados.items():
    print(f'{nome}: {item_encontrado}')