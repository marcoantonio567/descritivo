from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import messagebox, ttk ,filedialog
from decimal import Decimal, InvalidOperation
from PIL import Image as PILImage, ImageOps
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from num2words import num2words
from datetime import datetime
from docx.shared import Pt
from docx import Document
import win32com.client
from copy import copy
import tkinter as tk
import shutil
import os
import re


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
    workbook = load_workbook(caminho_arquivo, data_only=True)
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
            workbook = load_workbook(arquivo)
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
        rotulo_instrucao = ttk.Label(janela, text="Selecione as declividades abaixo:", background="#e6f7ff", font=("Helvetica", 12, "bold"))
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
            workbook = load_workbook(arquivo)
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
        rotulo_instrucao = ttk.Label(janela, text="Selecione as pedologias abaixo:", background="#e6f7ff", font=("Helvetica", 12, "bold"))
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
def substituir_cabecalho(texto,entrada,saida):
    # Carregar o documento Word
    doc = Document(entrada)

    # Acessar o cabeçalho da seção 1
    section = doc.sections[0]
    header = section.header

    # Modificar apenas o texto do cabeçalho, preservando imagens e formatação
    for paragraph in header.paragraphs:
        if paragraph.text.strip():
            paragraph.clear()
            run = paragraph.add_run(texto)
            run.font.size = Pt(11)
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            break

    # Salvar o documento modificado
    doc.save(saida)

    print("O cabeçalho da Seção 1 foi modificado com sucesso.")
def extrair_nomes_pedologias(pedologias):
    return [nome.split(' (')[0] for nome, descricao in pedologias]
def fazer_texto_pedologia(lista_pedologias):
    
    if not lista_pedologias:
        raise ValueError("A lista de pedologias está vazia.")

    if len(lista_pedologias) == 1:
        pedologia = lista_pedologias[0]
        texto_atualizado = f"O solo predominante no imóvel é o {pedologia}"
    else:
        pedologias_formatadas = " e ".join(", ".join(lista_pedologias).rsplit(", ", 1))
        texto_atualizado = f"Os solos predominantes no imóvel são os {pedologias_formatadas}"

    return texto_atualizado
def extrair_iniciais_desclividades(data):

    initials = []
    for item in data:
        if isinstance(item, tuple) and len(item) > 0:
            first_element = item[0]
            if isinstance(first_element, str) and len(first_element) > 0:
                initials.append(first_element[0])
    return initials
def pegar_maximo_2_intes_da_lista(input_list):
    if len(input_list) > 2:
        return input_list[:2]
    return input_list
def fazer_Texto_mosaico(letras):
    if len(letras) == 1:
        texto = f'{letras[0]} – Mosaico com predomínio de {letras[0]}'
    else:
        texto = f'{letras[0]}{letras[1]} – Mosaico com predomínio de {letras[0]} sobre {letras[1]}'
    return texto
def fazer_titulo_Declividade(lst):
    # Verifica se a lista é vazia
    if not lst:
        return ""

    # Divide a lista em grupos de 2
    groups = [lst[i:i + 2] for i in range(0, len(lst), 2)]

    # Formata os grupos
    formatted_groups = []
    for group in groups:
        formatted_groups.append("".join(group))

    # Junta os grupos com o "e" ao final
    if len(formatted_groups) > 1:
        return ", ".join(formatted_groups[:-1]) + " e " + formatted_groups[-1]
    else:
        return formatted_groups[0]
def contar_paginas_docx():
    word = 'teste.docx'
    nome_arquivo = word
    # Verifica se o arquivo existe no diretório atual
    if not os.path.isfile(nome_arquivo):
        print(f"O arquivo '{nome_arquivo}' não foi encontrado no diretório atual.")
        return

    try:
        # Inicializa o Word sem interface visível e sem a interface de COM (que é mais rápida)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Torna o Word invisível
        word.DisplayAlerts = 0  # Desabilita alertas que podem desacelerar o processo
        
        # Abre o documento sem carregar outros elementos (como fontes, imagens, etc.)
        documento = word.Documents.Open(os.path.abspath(nome_arquivo), ReadOnly=True)

        # Obtém o número de páginas
        numero_paginas = documento.ComputeStatistics(2)  # 2 significa wdStatisticPages
        
        # Fecha o documento e o Word
        documento.Close(False)
        word.Quit()

        # Exibe o número de páginas
        return str(numero_paginas)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None
def numero_por_extenso(numero):
    unidades = [
        "zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
        "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"
    ]
    dezenas = [
        "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"
    ]
    centenas = [
        "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"
    ]

    if int(numero) < 20:
        return unidades[int(numero)]
    elif int(numero) < 100:
        dezena = int(numero) // 10
        unidade = int(numero) % 10
        if unidade == 0:
            return dezenas[dezena]
        else:
            return f"{dezenas[dezena]} e {unidades[unidade]}"
    elif int(numero) < 1000:
        centena = int(numero) // 100
        resto = int(numero) % 100
        if resto == 0:
            return centenas[centena]
        else:
            return f"{centenas[centena]} e {numero_por_extenso(resto)}"
    else:
        return "Número fora do intervalo suportado."    
def colocar_quantidade_de_paginas_laudo():
    #tratando aqui as informações a respeito dos numeros de paginas
    #tem que ser por ultimo porque ele vai contar depois de todas as alterações feita
    numero_paginas = contar_paginas_docx()#aqui é pra eu saber quantas paginas tem no laudo
    numero_paginas_extenso = numero_por_extenso(numero_paginas)#aqui é pra transferir esse valor pra extenso
    dados_quantidade_pagina = {
        'hd190':numero_paginas,
        '8dg1':numero_paginas_extenso,
    }
    saida = 'teste.docx'
    substituir_palavras_documento(saida,dados_quantidade_pagina,saida)
def copiar_pagina_excel(destino):
    origem = "integracao.xlsx"
    nome_pagina = "quadro_resumo"

    # Carregar o arquivo de origem
    wb_origem = load_workbook(origem)
    if nome_pagina not in wb_origem.sheetnames:
        raise ValueError(f"A página '{nome_pagina}' não existe no arquivo de origem.")

    # Selecionar a página de origem
    pagina_origem = wb_origem[nome_pagina]

    # Carregar ou criar o arquivo de destino
    try:
        wb_destino = load_workbook(destino)
    except FileNotFoundError:
        wb_destino = load_workbook()

    # Criar uma nova página no arquivo de destino com o mesmo nome
    if nome_pagina in wb_destino.sheetnames:
        raise ValueError(f"A página '{nome_pagina}' já existe no arquivo de destino.")

    pagina_destino = wb_destino.create_sheet(nome_pagina)

    # Copiar os dados e a formatação célula por célula
    for linha in pagina_origem.iter_rows():
        for celula in linha:
            nova_celula = pagina_destino[celula.coordinate]
            nova_celula.value = celula.value

            if celula.has_style:
                nova_celula.font = copy(celula.font)
                nova_celula.border = copy(celula.border)
                nova_celula.fill = copy(celula.fill)
                nova_celula.number_format = celula.number_format
                nova_celula.protection = copy(celula.protection)
                nova_celula.alignment = copy(celula.alignment)

    # Ajustar larguras das colunas
    for col_idx, col_dim in pagina_origem.column_dimensions.items():
        pagina_destino.column_dimensions[col_idx].width = 8.43  # Definir largura da coluna

    # Ajustar alturas das linhas
    for row_idx in range(1, pagina_origem.max_row + 1):
        pagina_destino.row_dimensions[row_idx].height = 15  # Definir altura da linha

    # Copiar as configurações gerais da página
    pagina_destino.sheet_format = pagina_origem.sheet_format
    pagina_destino.sheet_properties = pagina_origem.sheet_properties
    pagina_destino.merged_cells = pagina_origem.merged_cells

    # Copiar imagens e manter tamanho e posição
    if hasattr(pagina_origem, '_images'):  # Verificar se a página contém imagens
        for img in pagina_origem._images:
            nova_imagem = Image(img.ref)
            nova_imagem.width = 585  # Definir a largura para 585
            nova_imagem.height = 400  # Definir a altura para 400
            pagina_destino.add_image(nova_imagem, img.anchor)

    # Salvar o arquivo de destino
    wb_destino.save(destino)
    print(f"Página '{nome_pagina}' copiada com sucesso de '{origem}' para '{destino}'.")
def selecionar_arquivo_excel():
    """Abre uma janela para o usuário selecionar um arquivo e retorna o caminho do arquivo."""
    # Cria uma janela oculta
    root = tk.Tk()
    root.withdraw()

    # Abre o seletor de arquivos e obtém o caminho do arquivo selecionado
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo que voce vai pegar as tabelas por favor",
        filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
    )

    if caminho_arquivo:
        print(f"Arquivo selecionado: {caminho_arquivo}")
    else:
        print("Nenhum arquivo foi selecionado.")

    return caminho_arquivo
def renovar_a_integração():#aqui vai ser preciso ser trocado para o pc do user
    origem = r'C:\\Users\\Usuario\\Desktop\\automatizar_descritivo\\TEMPLATES\\integracao.xlsx'
    destino = r'C:\\Users\\Usuario\\Desktop\\automatizar_descritivo\\'
    try:
        # Verifica se o arquivo de origem existe
        if not os.path.exists(origem):
            print(f"Arquivo de origem não encontrado: {origem}")
            return

        # Verifica se o destino é uma pasta
        if os.path.isdir(destino):
            destino = os.path.join(destino, os.path.basename(origem))

        # Copia o arquivo
        shutil.copy2(origem, destino)
        print(f"Arquivo copiado com sucesso para: {destino}")
    except Exception as e:
        print(f"Erro ao copiar o arquivo: {e}")
def selecionar_imagens_retornar_caminho(texto_cabecalho):
    # Cria uma janela oculta do Tkinter
    root = tk.Tk()
    root.withdraw()

    # Abre o seletor de arquivos e permite escolher múltiplas imagens
    caminhos = filedialog.askopenfilenames(
        title=texto_cabecalho,
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
def inserir_layout_geral_na_capa(image_path, cell):
    file_path = 'integracao.xlsx'
    width = 585  # largura da imagem
    height = 400  # altura da imagem
    sheet_name = 'quadro_resumo'

    # Criar o caminho para a imagem com borda no mesmo diretório
    image_dir, image_name = os.path.split(image_path)
    bordered_image_name = f"bordered_{image_name}"
    img_with_border_path = os.path.join(image_dir, bordered_image_name)

    # Adicionar borda leve à imagem
    try:
        with PILImage.open(image_path) as img:
            border_size = 2  # Tamanho da borda
            img_with_border = ImageOps.expand(img, border=border_size, fill="black")
            img_with_border.save(img_with_border_path)
    except Exception as e:
        print(f"Erro ao processar a imagem: {e}")
        return

    # Abrir o arquivo Excel
    try:
        workbook = load_workbook(file_path)
    except FileNotFoundError:
        print(f"Arquivo Excel '{file_path}' não encontrado.")
        return
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")
        return

    # Selecionar a página específica
    if sheet_name not in workbook.sheetnames:
        print(f"A página '{sheet_name}' não existe no arquivo.")
        return

    sheet = workbook[sheet_name]

    # Carregar a imagem com borda
    try:
        img = Image(img_with_border_path)
    except Exception as e:
        print(f"Erro ao carregar a imagem com borda: {e}")
        return

    # Redimensionar a imagem
    img.width = width
    img.height = height

    # Adicionar a imagem à célula especificada
    try:
        sheet.add_image(img, cell)
    except Exception as e:
        print(f"Erro ao adicionar a imagem na célula {cell}: {e}")
        return

    # Salvar as alterações
    try:
        workbook.save(file_path)
        print(f"Imagem inserida com sucesso na célula {cell} da página '{sheet_name}'.")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        return

    # Remover o arquivo temporário da imagem com borda
    try:
        os.remove(img_with_border_path)
        print(f"Imagem temporária '{img_with_border_path}' removida com sucesso.")
    except Exception as e:
        print(f"Erro ao remover o arquivo temporário: {e}")
def imagem_croqui():
    # Criar janela oculta para usar o seletor de arquivos
    root = tk.Tk()
    root.withdraw()  # Ocultar janela principal

    # Abrir seletor de arquivos
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a imagem do croqui de acesso",
        filetypes=[("Imagens PNG e JPG", "*.png;*.jpg"), ("Todos os Arquivos", "*.*")]
    )

    # Retornar o caminho do arquivo selecionado
    return caminho_arquivo
def arrumar_cpf_cnpj_proponente(caminho_estatistica):
    # Caminho para o arquivo Excel
    file_path = caminho_estatistica

    # Carrega o arquivo Excel
    workbook = load_workbook(file_path)

    # Seleciona a planilha (troque "Sheet1" pelo nome da sua planilha)
    sheet = workbook["quadro_resumo"]

    # Coordenada da célula mesclada (exemplo: A1)
    cell_address = "J2"

    # Verifica se a célula é mesclada
    for merged_cell in sheet.merged_cells.ranges:
        if cell_address in merged_cell:
            # Obtém o valor da célula mesclada
            cell_value = sheet[cell_address].value

            if cell_value:
                # Define o novo texto com base no número de caracteres
                if len(cell_value) == 20:
                    novo_texto = "CPF"
                else:
                    novo_texto = "CNPJ"

                # Substitui apenas o texto específico dentro do valor existente
                novo_valor = cell_value.replace("#486", novo_texto)
                sheet[cell_address] = novo_valor

            break

    # Salva as alterações
    workbook.save(file_path)

    print(f"O texto na célula {cell_address} foi atualizado com sucesso!")


























