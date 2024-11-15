from decimal import Decimal, InvalidOperation
from docx import Document
from datetime import datetime
from num2words import num2words
import openpyxl
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




