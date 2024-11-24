from docx import Document

def inserir_linha_abaixo(documento_path, palavra_alvo):
    # Carrega o documento
    doc = Document(documento_path)

    # Percorre todas as tabelas do documento
    for tabela in doc.tables:
        for linha_idx, linha in enumerate(tabela.rows):
            for celula in linha.cells:
                if palavra_alvo in celula.text:
                    # Cria uma nova linha abaixo da linha atual
                    nova_linha = tabela.add_row()
                    # Copia o estilo da linha anterior
                    for idx, cel in enumerate(nova_linha.cells):
                        cel.text = ''
                    # Insere o texto na primeira célula da nova linha
                    nova_linha.cells[0].text = 'Texto inserido abaixo da palavra alvo'

    # Salva o documento com a linha inserida
    doc.save('documento_atualizado.docx')

# Caminho para o documento
caminho_documento = 'documento_atualizado.docx'
# Palavra que você deseja encontrar
palavra = '#486'

inserir_linha_abaixo(caminho_documento, palavra)
