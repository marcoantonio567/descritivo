import xlwings as xw
import pyperclip

# Função para copiar dados e formatação de uma planilha para a área de transferência
def copiar_dados_com_formatacao(caminho_arquivo, nome_planilha, intervalo):
    # Inicia o Excel
    app = xw.App(visible=False)  # O Excel não precisa ser visível
    wb = app.books.open(caminho_arquivo)
    
    # Seleciona a planilha
    sheet = wb.sheets[nome_planilha]
    
    # Seleciona o intervalo de células
    intervalo_celulas = sheet.range(intervalo)
    
    # Copia o intervalo com formatação
    intervalo_celulas.api.Copy()  # Usando a API do Excel para copiar com formatação
    
    # Fecha o Excel sem salvar alterações
    wb.close()
    app.quit()
    
    print(f"Dados e formatação da planilha '{nome_planilha}' copiados para a área de transferência!")
# Exemplo de uso
caminho_arquivo = 'integracao.xlsx'  # Substitua pelo caminho do seu arquivo Excel
nome_planilha = 'quadro_resumo'  # Nome da planilha desejada
intervalo = 'A1:N20'  # Intervalo de células que você quer copiar
copiar_dados_com_formatacao(caminho_arquivo, nome_planilha, intervalo)
 