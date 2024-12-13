import pandas as pd
import re
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client

def excel_col_to_index(col_str):
    """Converte o rótulo de coluna do Excel (por ex: 'A', 'B', 'C', 'AA') para índice baseado em zero."""
    result = 0
    for char in col_str:
        result = result * 26 + (ord(char.upper()) - ord('A')) + 1
    return result - 1  # zero-based index

def parse_cell_address(cell_addr):
    """Recebe um endereço de célula Excel (ex: 'C3') e retorna (row_index, col_index) zero-based."""
    match = re.match(r"([A-Za-z]+)([0-9]+)", cell_addr)
    if not match:
        raise ValueError(f"Endereço de célula inválido: {cell_addr}")
    col_str = match.group(1)
    row_str = match.group(2)
    col_index = excel_col_to_index(col_str)
    row_index = int(row_str) - 1  # zero-based
    return (row_index, col_index)

def format_number(value, decimal_places, replace_dot_with_comma=True):
    if pd.isnull(value):
        return ""
    # Tenta converter para float
    str_value = str(value).replace(',', '.')
    try:
        num = float(str_value)
    except ValueError:
        # Se não conseguir converter para número, retorna o valor original
        return str(value)

    formatted = f"{num:.{decimal_places}f}"
    if replace_dot_with_comma:
        formatted = formatted.replace('.', ',')
    return formatted

def format_currency(value):
    if pd.isnull(value):
        return ""
    str_value = str(value).replace(',', '.')
    try:
        num = float(str_value)
    except ValueError:
        return str(value)
    # Formata primeiro no padrão US: "9,297.52"
    us_formatted = f"{num:,.2f}"
    # Agora converte para padrão brasileiro
    temp = us_formatted.replace('.', '_').replace(',', '.').replace('_', ',')
    formatted = f"R$ {temp}"
    return formatted

def format_phone(value):
    if pd.isnull(value):
        return ""
    phone_str = str(value).strip()
    # Se já tiver parênteses, assumimos que já possui DDD formatado
    if phone_str.startswith('('):
        return phone_str
    else:
        return f"(11) {phone_str}"

def extrair_intervalo_excel_por_celulas(arquivo, planilha, intervalo_str):
    # intervalo_str no formato 'C3:L9'
    start_cell, end_cell = intervalo_str.split(':')
    start_row, start_col = parse_cell_address(start_cell)
    end_row, end_col = parse_cell_address(end_cell)

    df = pd.read_excel(arquivo, sheet_name=planilha, header=None)

    df_intervalo = df.iloc[start_row:end_row+1, start_col:end_col+1]

    # Coluna 1: 4 casas decimais, vírgula
    if df_intervalo.shape[1] >= 1:
        df_intervalo.iloc[:,0] = df_intervalo.iloc[:,0].apply(lambda x: format_number(x, 4, True))

    # Coluna 2 e 3: moeda
    if df_intervalo.shape[1] >= 2:
        df_intervalo.iloc[:,1] = df_intervalo.iloc[:,1].apply(format_currency)
    if df_intervalo.shape[1] >= 3:
        df_intervalo.iloc[:,2] = df_intervalo.iloc[:,2].apply(format_currency)

    # Coluna 4: até 4 casas decimais, vírgula
    if df_intervalo.shape[1] >= 4:
        df_intervalo.iloc[:,3] = df_intervalo.iloc[:,3].apply(lambda x: format_number(x, 4, True))

    # Colunas 5 a 8: 2 casas decimais, vírgula
    for col_idx in range(4, 8):
        if df_intervalo.shape[1] > col_idx:
            df_intervalo.iloc[:,col_idx] = df_intervalo.iloc[:,col_idx].apply(lambda x: format_number(x, 2, True))

    # Coluna 9: texto normal (sem alteração)
    # Coluna 10: telefone
    if df_intervalo.shape[1] >= 10:
        df_intervalo.iloc[:,9] = df_intervalo.iloc[:,9].apply(format_phone)

    # Retorna o DataFrame
    return df_intervalo

def fazer_lista_codigos_quadro():
    numero_incio = 10
    codigo_inicio = f'%code'
    lista_codigos = []
    while numero_incio <= 69:
        lista_codigos.append(f'{codigo_inicio}{numero_incio}')
        numero_incio +=1
    return lista_codigos

def substituir_palavras_no_word(caminho_arquivo, caminho_arquivo_saida, substituicoes):

    try:
        # Inicia a aplicação do Word via COM
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Executar em segundo plano

        # Abre o documento
        doc = word.Documents.Open(caminho_arquivo)

        # Para cada chave e valor, realizar a substituição
        for chave, valor in substituicoes.items():
            # Configurando o objeto Find
            find = doc.Content.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()

            find.Text = chave
            find.Replacement.Text = valor
            
            # Configurações adicionais
            find.Forward = True
            find.Wrap = 1  # wdFindContinue
            find.MatchCase = False
            find.MatchWholeWord = True
            find.MatchWildcards = False
            find.MatchSoundsLike = False
            find.MatchAllWordForms = False

            # wdReplaceAll = 2
            resultado = find.Execute(Replace=2)
            if not resultado:
                print(f"A substituição para '{chave}' não foi realizada.")

        # Salva e fecha o documento
        doc.SaveAs2(caminho_arquivo_saida)
        doc.Close()
        word.Quit()

        print("Substituições concluídas com sucesso!")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        if 'word' in locals():
            word.Quit()

def criar_dicionario(chave,dado):
   
    if len(chave) != len(dado):
        raise ValueError("As listas devem ter o mesmo tamanho.")

    return dict(zip(chave, dado))

caaminho_laudo = r'C:\Users\Usuario\Desktop\automatizar_descritivo\LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
df_resultado = extrair_intervalo_excel_por_celulas("estattis_automação.xlsx", "QUADRO", "C4:L9")
lista_codgios_quadro = fazer_lista_codigos_quadro()
lista_dados = [item for sublista in df_resultado.values.tolist() for item in sublista] # Criando uma única lista com todos os elementos
dicionario = criar_dicionario(lista_codgios_quadro,lista_dados)
substituir_palavras_no_word(caaminho_laudo,'teste.docx',dicionario)