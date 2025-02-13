from colar_imagens import colar_maps ,adcionar_imagens_documentos ,delete_last_page , substituir_croqui_e_rota_acesso
from leitura_excel import *
from funcoes import *
from textos import *
from cores import *
from Ia import perguntar_groq

#aqui é pra ele fazer o cabeçalho
word_filee = 'teste.docx'
substituir_cabecalho(texto_cabecalho,word_filee,word_filee)


#tratando as datas
data_solicitacao_limpo = formatar_data(data_solicitacao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_emissao_limpo = formatar_data(data_emissao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_vistoria_limpo = formatar_data(data_vistoria)#aqui to to tratando a data pra ela ficar assim : 00/00/0000


#colocando virgulas nas areas ao inves de pontos
area_ha_limpo = substituir_ponto_por_virgulas(area_ha,casas_decimais=4)
area_construida_limpo = substituir_ponto_por_virgulas(area_construida,casas_decimais=2)
area1_limpo = substituir_ponto_por_virgulas(area1,casas_decimais=4)
area2_limpo = substituir_ponto_por_virgulas(area2,casas_decimais=4)



#chamando funções
data_atual = gerar_data_atual()


declividade_vegetação = escolha_usuario()
if declividade_vegetação == 'declividade':
    declividades = selecionar_declividade()#lista de declividades e suas descrições
    texto_das_desclividades = gerar_texto(declividades)#aqui vai gerar um texto com todas as declividades separando elas por paaragrafo
    inciais_declividades = extrair_iniciais_desclividades(declividades)#aqui ele vai pegar as inicias das declividades
    texo_mosaico = escolher_Texto_mosaicos()#aqui ele vai gerar um texto que vai pro mosaico
    if texo_mosaico is None:
        texo_mosaico = ""
    titulo_declividades = fazer_titulo_Declividade(inciais_declividades)#aqui ele vai fazer o titulo da declividade
    resposta_introducao_declividade = texto_declividades_titulo(inciais_declividades,quantidade_matriculas,titulo_declividades)
if declividade_vegetação == 'vegetação': #aqui eu vou colcoar as variaveis como declividade_porque pra não precisar tratar mais dados
    declividades = selecionar_vegetacao()
    texto_das_desclividades = gerar_texto(declividades)#aqui vai gerar um texto com todas as vegetações separando elas por paaragrafo
    texo_mosaico = ""
    nomes_vegetacao = extrair_nomes_vegetacao(declividades)
    titulo_declividades = fazer_titulo_Declividade(nomes_vegetacao)#aqui ele vai fazer o titulo da vegetação
    resposta_introducao_declividade = texto_vegetacao_titulos(nomes_vegetacao,quantidade_matriculas,titulo_declividades)


pedologias = selecionar_pedologia()#lista de pedologias e suas descrições
texto_das_pedologias = gerar_texto(pedologias)#aqui vai gerar um texto com todas as pedologia separando elas por paaragrafo
nomes_pedologias = extrair_nomes_pedologias(pedologias)#aqui eu ele vai pegar apenas os nomes das pedologias da lista de pedologias selecionadas pelo usuario
titulo_pedologias = fazer_texto_pedologia(nomes_pedologias)#aqui ele vai fazer o texto do titulo das pedologias


#tratando se é cpf ou cnpj 
#propietario
if genero == 'empresa' or len(cpf_cpnj_propietario) == 18:
    reposta_cpf_cnpj_propietario = "CNPJ"
else:
    reposta_cpf_cnpj_propietario = "CPF"

#proponente
if genero == 'empresa' or len(cpf_cpnj_proponente) == 18:
    reposta_cpf_cnpj_proponente = "CNPJ"
else:
    reposta_cpf_cnpj_proponente = "CPF"

#arrumando bacia e sub-bacia
def resposta_sistema_hidrografico():
    return sistema_hidrografico()

#resposta para o memorial descritivo
def fazer_texto_georeferenciamento():
    resposta_gel_referenciamento = []
    for item in data_imoveis:
        if georeferenciamento == 'GEO':
            resposta_gel_referenciamento.append(texto_do_gel(str(item[0])))
        elif georeferenciamento == 'CAR':
            resposta_gel_referenciamento.append(texto_car(str(item[0])))
        elif georeferenciamento == 'Memorial descritivo':
            resposta_gel_referenciamento.append(texto_memorial_descritivo(str(item[0])))
        elif georeferenciamento == 'CAR e Memorial':
            resposta_gel_referenciamento.append(texto_memorial_descritivo_com_car(str(item[0])))
        elif georeferenciamento == 'CAR e GEO':
            resposta_gel_referenciamento.append(texot_car_geo(str(item[0])))
        
    texto_output = "\n\n".join(resposta_gel_referenciamento)
    return texto_output

def registro_imovel():
    return texto_registro(quantidade_matriculas)

#verificando se o imovel tem hipotecas pendentes ou não
def hipotecas_reposta():
    return texto_hipotecas(quantidade_matriculas)

#verificação se o imovel esta inserido no bioma amazonico
def reposta_bioma():
    return texto_bioma(quantidade_matriculas)


def resposta_texto_do_final_liquidacao_forcada():
    return texto_do_final_da_liquidacao_forcada()

#aqui eu to copiando a tabela que vai na capa para a tabela de estatistica
#TODO: trocar aqui pelo do user
diretorios_base = [
        r"Z:\\1. AVALIAÇÕES\\01. AVALIAÇÕES SICREDI\\01. RURAL\\",
        r"Z:\\1. AVALIAÇÕES\\02. AVALIAÇÕES SICOOB\\01. RURAL\\",
        r"Z:\\1. AVALIAÇÕES\\03. SICOOB CREDICOM\\",
        r"Z:\\1. AVALIAÇÕES\\04. SICOOB ENGECRED\\",
        r"Z:\\1. AVALIAÇÕES\\05. AVALIAÇÕES BRB\\",
        r"Z:\\1. AVALIAÇÕES\\06. AVALIAÇÕES CAIXA\\",
        r"Z:\\1. AVALIAÇÕES\\08. PARTICULAR\\"
    ]
destino_excel_estatistica = encontrar_primeiro_excel_pecas(diretorios_base,fluid)




#aqui eu to indo selecionar os mapas
lista_de_imagens_mapas = listar_imagens_na_pasta_mapas(destino_excel_estatistica)
#aqui é a lista de mapas  que vai no laudo
nomes_procurados_mapas = ["layout geral", "PEDOLOGIA", "vegetação", "bacia" ,"declividade"]
#auqi são os resultados em chaves
caminhos_dos_mapas = encontrar_nomes(lista_de_imagens_mapas,nomes_procurados_mapas)
lista_de_caminhos_mapas_limpo = list(caminhos_dos_mapas.values())#aqui eu to transformando as chaves em lista para eu conseguir seleiconar eles
#aqui abaixo eu to puxando todos os caminhos das imagens
layout_geral_mapa = lista_de_caminhos_mapas_limpo[0] if len(lista_de_caminhos_mapas_limpo) > 0 else None
pedologia_mapa = lista_de_caminhos_mapas_limpo[1] if len(lista_de_caminhos_mapas_limpo) > 1 else None
vegetacao_mapa = lista_de_caminhos_mapas_limpo[2] if len(lista_de_caminhos_mapas_limpo) > 2 else None
bacia_ou_sub_bacia_mapa = lista_de_caminhos_mapas_limpo[3] if len(lista_de_caminhos_mapas_limpo) > 3 else None
declividade_mapa = lista_de_caminhos_mapas_limpo[4] if len(lista_de_caminhos_mapas_limpo) > 4 else None

if declividade_vegetação == 'vegetação': #aqui é  pra ele colocar a vegetação
    declividade_mapa = vegetacao_mapa



rota_de_acesso = escolher_e_ler_arquivo_txt(destino_excel_estatistica)#escolher rota de acesso atraves da estatistica
croqui_imagem = encontrar_croqui_no_diretorio(destino_excel_estatistica)#escolher o croqui atraves da estatistica
substituir_croqui_e_rota_acesso(word_filee,'{{dh19a}}', croqui_imagem, rota_de_acesso, width_cm=15.0, height_cm=8.77, border_size=3, border_color="black")#colocando a rota e o croqui de acesso indepedente de quantos sejam

pasta_pecas_tecnicas = caminho_pasta_pecas_tecnicas(destino_excel_estatistica)
images_with_placeholders = [#essa variavel aqui contem uma lista com os caminhos dos mapas
    (pedologia_mapa, "{{h01hf}}"),
    (bacia_ou_sub_bacia_mapa, "{{h9fd1}}"),
    (declividade_mapa, "{{g7aa}}"),
]

#Insere as imagens substituindo os marcadores de posição
colar_maps(word_filee, images_with_placeholders)


caminho_imagens_cit = enontrar_imagens_documentos(destino_excel_estatistica,'cit')
caminho_imagens_ccir = enontrar_imagens_documentos(destino_excel_estatistica,'ccir')
caminho_imagens_car = enontrar_imagens_documentos(destino_excel_estatistica,'car')
caminho_imagens_cnd = enontrar_imagens_documentos(destino_excel_estatistica,'cnd')
if caminho_imagens_cnd is None:#aqui eu to falando pra se ele não encontrar a cnd ele procure direto o itr
    caminho_imagens_cnd = enontrar_imagens_documentos(destino_excel_estatistica,'itr')

lista_geral_imagens = [caminho_imagens_cit,caminho_imagens_cnd,caminho_imagens_ccir,caminho_imagens_car]
quantidade_de_imagens = sum(len(lista) for lista in lista_geral_imagens)
print(f'{DARK_GREEN} a quantidade de imagens é :{quantidade_de_imagens}{RESET}')

adcionar_imagens_documentos(word_filee,caminho_imagens_cit,'Documentos anexados\nCERTIDÃO DE INTEIRO TEOR DA MATRÍCULA')
adcionar_imagens_documentos(word_filee,caminho_imagens_car,'RECIBO DE INSCRIÇÃO DO IMÓVEL RURAL NO CAR')
adcionar_imagens_documentos(word_filee,caminho_imagens_cnd,'CERTIDÃO NEGATIVA DE DÉBITOS – CND')
adcionar_imagens_documentos(word_filee,caminho_imagens_ccir,'CERTIFICADO DE CADASTRO DE IMÓVEL RURAL – CCIR')
delete_last_page(word_filee)#aqui vou apagar a ultima pagina porque como ele quebra todas as paginas então ele adcionar uma pagina a mais por isso to tirando a ultima

#arrumando pra ele pegar o texto coeficiente de variação
estatisticas_laudos = enontrar_estatisticas(caminho_imagens_cit[0])#aqui é pra ele pegar o primeiro item da kista ja que o caminho é uma lista
lista_numeros_matriculas = []
valores_coeficientes = []
municipios = []
estado_uf = []
valores_mercado_automaticos = [] # oque eu to fazendo aqui é colocando o valor de mercado para se puxar de forma automatica em cada estatistica
valores_liquidacao_automaticos = [] # oque eu to fazendo aqui é colocando o valor de liquidacao para se puxar de forma automatica em cada estatistica
for estatistica in estatisticas_laudos:
    texto_da_matricula = buscar_valor_excel(estatistica,"SANEAMENTO", "C4")
    numeros_da_matricula = extrair_numeros(texto_da_matricula)
    lista_numeros_matriculas.append(numeros_da_matricula)
    valor_coeficiente = buscar_valor_excel(estatistica,"PLANILHA HOMOG", "Q24")
    if valor_coeficiente is not None:
        valor_coeficiente_formatado = f"{float(valor_coeficiente):.2f}" #aqui é pra ela formatar o valor do coeficiente no maximo duas casas decimais
    if valor_coeficiente is None:
        print(f'{DARK_RED} o valor de coeficiente da matricula não foi encontrado{RESET}')
        valor_coeficiente_formatado = 'nofind'
    valores_coeficientes.append(valor_coeficiente_formatado)
    
    cidade_estado_junto = buscar_valor_excel(estatistica,"AMOSTRAS","E5")
    municipio , uf = extrair_cidade_uf(cidade_estado_junto)
    municipios.append(municipio)
    estado_uf.append(uf)

    valor_mercado_automatico = buscar_valor_excel(estatistica,'SANEAMENTO','F20')
    valor_mercado_automatico_limpo = formatar_valor(valor_mercado_automatico)
    valor_mercado_arredondado = ajustar_milhar(valor_mercado_automatico_limpo)
    valores_mercado_automaticos.append(valor_mercado_arredondado)

    valor_liquidacao_automatico = buscar_valor_excel(estatistica,'LIQUIDAÇÃO','F11')
    
    valor_liquidacao_automatico_limpo = formatar_valor(valor_liquidacao_automatico)
    valor_liquidacao_arredondado = ajustar_milhar(valor_liquidacao_automatico_limpo)
    valores_liquidacao_automaticos.append(valor_liquidacao_arredondado)

texto_coeficientes = gerar_texto_coeficientes(valores_coeficientes,lista_numeros_matriculas)

#aqui a baixo eu to redefinindo os itens 4 e 5 da minha lista de imoveis
data_imoveis = substituir_indices_4_e_5(data_imoveis,valores_mercado_automaticos,valores_liquidacao_automaticos)

#convertendo os valores para moeda
valor_mercado_limpo = valores_mercado_automaticos[0]
liquidacao_forcada_limpo = valores_liquidacao_automaticos[0]

def reposta_valores_de_mercado():
    return valores_mercado(data_imoveis)

def resposta_valores_liquidacao_forcada():
    return valores_liquidacao_forcada(data_imoveis)

#aqui eu to fazendo o texto do valor de mercado do item 14
def resposta_texto_mercado():
    return texto_valor_mercado(municipios[0],estado_uf[0])

def resposta_texto_data_emissao():
    return texto_data_emissao()

def respota_da_ia():
    pergunta_pra_ia = f'me fale sobre o solo/clima/vegetação/historia da cidade {municipios[0]} - {estado_uf[0]}'
    respostaa_Da_pergunta = perguntar_groq(pergunta_pra_ia)
    print(respostaa_Da_pergunta)
    return respostaa_Da_pergunta
