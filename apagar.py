import win32com.client
import os

def contar_paginas_word():
    caminho_arquivo = 'LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
    if not os.path.exists(caminho_arquivo):
        print("Arquivo não encontrado.")
        return None
    
    try:
        # Inicializa o Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.ScreenUpdating = False  # Evita atualizações de tela

        # Converte o caminho para absoluto
        caminho_absoluto = os.path.abspath(caminho_arquivo)

        # Abre o documento
        doc = word.Documents.Open(caminho_absoluto)

        try:
            # Obtém o número de páginas
            num_paginas = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
        finally:
            # Fecha o documento
            doc.Close(SaveChanges=False)
        
        return num_paginas
    except Exception as e:
        print(f"Erro ao processar o documento: {e}")
        return None
    finally:
        # Fecha o Word
        word.Quit()
    
# Exemplo de uso
print(f"O documento tem {contar_paginas_word()} páginas.")
