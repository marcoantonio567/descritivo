import os
from docx import Document
from PIL import ImageGrab
import win32com.client as win32
import win32clipboard as clipboard
from docx.shared import Inches
import pandas as pd
import tkinter as tk
from tkinter import filedialog



def buscar_excel():
    # Cria uma janela oculta do tkinter
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    
    # Abre um diálogo para selecionar um arquivo Excel
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    # Verifica se um arquivo foi selecionado
    if caminho_arquivo:
        print(f"Caminho do arquivo selecionado: {caminho_arquivo}")
        return caminho_arquivo
    else:
        print("Nenhum arquivo selecionado.")
        return None

def colar_Tabelas(pagina_xlsx,intervalo_tabela,code_substituicao,tamanho):
    file_docx= 'LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
    excel_file = documento_Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Manter o Excel invisível
    workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
    sheet = workbook.Sheets[pagina_xlsx]

    # Definir o intervalo da segunda tabela manualmente
    # Suponha que a segunda tabela esteja em 'B15:E25' (este é um exemplo; você pode alterar conforme necessário)
    second_table_range = sheet.Range(intervalo_tabela)
    second_table_range.Copy()

    # Capturar a imagem da tabela usando o clipboard
    img = ImageGrab.grabclipboard()
    if img is not None:
        img_path = "tabela_imagem_temp.png"  # Salvar a imagem temporariamente
        img.save(img_path, 'PNG')
    else:
        raise ValueError("Erro ao copiar a imagem da tabela para o clipboard.")

    # Limpar a área de transferência para liberar a memória
    clipboard.OpenClipboard()
    clipboard.EmptyClipboard()
    clipboard.CloseClipboard()

    # Fechar o Excel
    workbook.Close(False)
    excel.Quit()

    # Passo 2: Abrir o documento Word existente e substituir o placeholder pela imagem
    document_path = file_docx  # Documento que já possui o texto placeholder
    doc = Document(document_path)

    # Procurar o placeholder e substituir pela imagem
    placeholder = code_substituicao

    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # Limpar o texto do placeholder
            paragraph.text = paragraph.text.replace(placeholder, "")
            # Inserir a imagem após o parágrafo
            run = paragraph.add_run()
            run.add_picture(img_path, width=Inches(tamanho))  # Defina a largura como desejar

    # Salvar o documento Word
    document_path_output = 'LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
    doc.save(document_path_output)

    # Remover a imagem temporária
    os.remove(img_path)

    print("alteração feito com sucesso")


documento_Excel = 'estattis_automação.xlsx'
#area ultil do imove
colar_Tabelas('AREA UTIL ','C4:J13','#sdkjbf',tamanho=6.8)
#area de uso do imovel
colar_Tabelas('AREA UTIL ','C16:I20','#5213',tamanho=5.9)
#quadro de homogeneização
colar_Tabelas('PLANILHA HOMOG','A3:R25','#12863',tamanho=10.4)
#saneamento das amostras
colar_Tabelas('SANEAMENTO','C4:H14','#28731',tamanho=3.9)
#valor da terra nua apurado
colar_Tabelas('SANEAMENTO','C16:H20','#123452',tamanho=4)
#calculo do valor de liquidação forçada 
colar_Tabelas('LIQUIDAÇÃO','C5:H11','#1313',tamanho=4.5)
#quadro de amostras
colar_Tabelas('QUADRO','B3:L9','#12746',tamanho=9.7)