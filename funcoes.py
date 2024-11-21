from decimal import Decimal, InvalidOperation
from docx import Document
from datetime import datetime
from num2words import num2words
import openpyxl
import os
import re
import tkinter as tk
from tkinter import messagebox, ttk ,filedialog

def formatar_data(data_str):
    data = datetime.strptime(data_str, "%Y-%m-%d %H:%M:%S")
    return data.strftime("%d/%m/%Y")
def substituir_palavras_documento(doc_path, substituicoes, output_path):
    documento = Document(doc_path)
    
    for paragrafo in documento.paragraphs:
        for run in paragrafo.runs:
            for palavra_antiga, palavra_nova in substituicoes.items():
                if palavra_antiga in run.text:
                    run.text = run.text.replace(palavra_antiga, palavra_nova)

    for tabela in documento.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for run in paragrafo.runs:
                        for palavra_antiga, palavra_nova in substituicoes.items():
                            if palavra_antiga in run.text:
                                run.text = run.text.replace(palavra_antiga, palavra_nova)

    documento.save(output_path)
    print(f"Substituição concluída. Documento salvo em {output_path}")
def gerar_data_atual():
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
             "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    
    data = datetime.now()
    dia = data.day
    mes = meses[data.month - 1]
    ano = data.year
    
    return f"Palmas - TO, {dia} de {mes} de {ano}."
def ler_celula_excel(celula):
    caminho_arquivo = 'integracao.xlsx'
    nome_planilha = 'dados' 
    # Abre o arquivo Excel com data_only=True para ler o valor das fórmulas
    workbook = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    # Seleciona a planilha especificada
    planilha = workbook[nome_planilha]
    # Obtém o valor da célula especificada
    valor = planilha[celula].value
    return valor
def formatar_valor(valor):
    valor = float(valor)
    valor_formatado = "R$ {:,.2f}".format(valor)
    valor_formatado = valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")
    return valor_formatado
def abrir_arquivo_word(caminho_arquivo):
    try:
        # Abre o arquivo Word no aplicativo padrão do sistema
        os.startfile(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")
def valor_por_extenso(valor_str):
    # Remover o 'R$' e os pontos, substituir a vírgula por ponto para converter em float
    valor_numerico = float(re.sub(r'[^0-9,]', '', valor_str).replace(',', '.'))
    
    # Parte inteira e parte decimal
    parte_inteira = int(valor_numerico)
    parte_decimal = int(round((valor_numerico - parte_inteira) * 100))

    # Converter parte inteira e decimal para extenso
    extenso_inteira = num2words(parte_inteira, lang='pt_BR')
    extenso_decimal = num2words(parte_decimal, lang='pt_BR')
    
    # Montar a resposta por extenso
    if parte_decimal > 0:
        valor_extenso = f"{extenso_inteira} reais e {extenso_decimal} centavos"
    else:
        valor_extenso = f"{extenso_inteira} reais"

    return valor_extenso
def substituir_ponto_por_virgulas(varivel):
    return varivel.replace(".",",")
def escolher_e_ler_arquivo_txt():
    # Cria uma interface gráfica oculta apenas para usar o filedialog
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal

    # Abre a janela para o usuário escolher um arquivo
    caminho_arquivo = filedialog.askopenfilename(
        title="Escolha um arquivo de bloco de notas",
        filetypes=[("Text files", "*.txt")]
    )

    if caminho_arquivo:
        try:
            with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
                conteudo = arquivo.read()
                
                return conteudo
        except FileNotFoundError:
            print("Arquivo não encontrado.")
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
    else:
        return None
def selecionar_declividade():
    arquivo = 'integracao.xlsx'
    aba_nome = 'DECLIVIDADE E PEDOLOGIA'  # Nome da aba onde os dados estão localizados
    coluna = 2  # Número da coluna que você quer ler (por exemplo, 2 para coluna B)

    # Função para ler a coluna específica da aba
    def ler_coluna_excel(arquivo, aba_nome, coluna):
        try:
            workbook = openpyxl.load_workbook(arquivo)
            aba = workbook[aba_nome]
            valores = []
            for linha in aba.iter_rows(min_col=coluna, max_col=coluna + 1, min_row=3,max_row=8, values_only=True):
                if linha[0] is not None:
                    valores.append((linha[0], linha[1]))  # Adiciona o valor da coluna e a célula da direita
            return valores
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{arquivo}' não encontrado.")
        except KeyError:
            messagebox.showerror("Erro", f"Aba '{aba_nome}' não encontrada no arquivo Excel.")
        return []

    # Função para exibir a lista e permitir que o usuário selecione um ou mais valores
    def selecionar_valor(valores):
        selecoes = []

        def confirmar_selecao():
            selecao = listbox.curselection()
            if selecao:
                for index in selecao:
                    selecionado = valores[index]
                    selecoes.append(selecionado)
                janela.destroy()
            else:
                messagebox.showwarning("Atenção", "Nenhum valor selecionado!")

        # Criação da interface gráfica
        janela = tk.Tk()
        janela.title("Seleção de Valores do Excel")
        janela.geometry("500x400")
        janela.configure(bg="#e6f7ff")

        # Rótulo de instrução
        rotulo_instrucao = ttk.Label(janela, text="Selecione um ou mais valores da lista abaixo:", background="#e6f7ff", font=("Helvetica", 12, "bold"))
        rotulo_instrucao.pack(pady=(10, 5))

        # Listbox com barra de rolagem
        frame_listbox = tk.Frame(janela, bg="#e6f7ff")
        frame_listbox.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL)
        listbox = tk.Listbox(frame_listbox, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, height=10, width=50, bg="#ffffff", fg="#003366", font=("Arial", 10))
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for valor in valores:
            listbox.insert(tk.END, valor[0])

        # Botão para confirmar a seleção
        botao_confirmar = ttk.Button(janela, text="Confirmar Seleção", command=confirmar_selecao)
        botao_confirmar.pack(pady=(5, 15))

        # Tornar a janela responsiva
        janela.grid_rowconfigure(0, weight=1)
        janela.grid_columnconfigure(0, weight=1)
        frame_listbox.grid_rowconfigure(0, weight=1)
        frame_listbox.grid_columnconfigure(0, weight=1)

        janela.mainloop()
        return selecoes

    # Lê a coluna do Excel e exibe a interface para seleção
    valores_coluna = ler_coluna_excel(arquivo, aba_nome, coluna)
    if valores_coluna:
        return selecionar_valor(valores_coluna)
    else:
        return []
def selecionar_pedologia():
    arquivo = 'integracao.xlsx'
    aba_nome = 'DECLIVIDADE E PEDOLOGIA'  # Nome da aba onde os dados estão localizados
    coluna = 2  # Número da coluna que você quer ler (por exemplo, 2 para coluna B)

    # Função para ler a coluna específica da aba
    def ler_coluna_excel(arquivo, aba_nome, coluna):
        try:
            workbook = openpyxl.load_workbook(arquivo)
            aba = workbook[aba_nome]
            valores = []
            for linha in aba.iter_rows(min_col=coluna, max_col=coluna + 1, min_row=14,max_row=33, values_only=True):
                if linha[0] is not None:
                    valores.append((linha[0], linha[1]))  # Adiciona o valor da coluna e a célula da direita
            return valores
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{arquivo}' não encontrado.")
        except KeyError:
            messagebox.showerror("Erro", f"Aba '{aba_nome}' não encontrada no arquivo Excel.")
        return []

    # Função para exibir a lista e permitir que o usuário selecione um ou mais valores
    def selecionar_valor(valores):
        selecoes = []

        def confirmar_selecao():
            selecao = listbox.curselection()
            if selecao:
                for index in selecao:
                    selecionado = valores[index]
                    selecoes.append(selecionado)
                janela.destroy()
            else:
                messagebox.showwarning("Atenção", "Nenhum valor selecionado!")

        # Criação da interface gráfica
        janela = tk.Tk()
        janela.title("Seleção de Valores do Excel")
        janela.geometry("500x400")
        janela.configure(bg="#e6f7ff")

        # Rótulo de instrução
        rotulo_instrucao = ttk.Label(janela, text="Selecione um ou mais valores da lista abaixo:", background="#e6f7ff", font=("Helvetica", 12, "bold"))
        rotulo_instrucao.pack(pady=(10, 5))

        # Listbox com barra de rolagem
        frame_listbox = tk.Frame(janela, bg="#e6f7ff")
        frame_listbox.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(frame_listbox, orient=tk.VERTICAL)
        listbox = tk.Listbox(frame_listbox, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, height=10, width=50, bg="#ffffff", fg="#003366", font=("Arial", 10))
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for valor in valores:
            listbox.insert(tk.END, valor[0])

        # Botão para confirmar a seleção
        botao_confirmar = ttk.Button(janela, text="Confirmar Seleção", command=confirmar_selecao)
        botao_confirmar.pack(pady=(5, 15))

        # Tornar a janela responsiva
        janela.grid_rowconfigure(0, weight=1)
        janela.grid_columnconfigure(0, weight=1)
        frame_listbox.grid_rowconfigure(0, weight=1)
        frame_listbox.grid_columnconfigure(0, weight=1)

        janela.mainloop()
        return selecoes

    # Lê a coluna do Excel e exibe a interface para seleção
    valores_coluna = ler_coluna_excel(arquivo, aba_nome, coluna)
    if valores_coluna:
        return selecionar_valor(valores_coluna)
    else:
        return []
def gerar_texto(lista_tuplas):
    texto = ""
    for item in lista_tuplas:
        paragrafo = f"\t{item[0]} {item[1]}\n"
        texto += paragrafo
    return texto













