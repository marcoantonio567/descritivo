from fazer_capa_excel import celula_para_colar_o_mapa #aqui quando eu to chamando essa variavel ja fazendo a capa do excel
from colar_imagens import colar_maps_e_croquii , colar_imagens_documentos
from leitura_excel import *
from funcoes import *
from textos import *


#aqui é pra ele fazer o cabeçalho
word_filee = 'LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
substituir_cabecalho(texto_cabecalho,word_filee,word_filee)


#tratando as datas
data_solicitacao_limpo = formatar_data(data_solicitacao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_emissao_limpo = formatar_data(data_emissao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_vistoria_limpo = formatar_data(data_vistoria)#aqui to to tratando a data pra ela ficar assim : 00/00/0000


#colocando virgulas nas areas ao inves de pontos
area_ha = substituir_ponto_por_virgulas(area_ha,casas_decimais=4)
area_construida = substituir_ponto_por_virgulas(area_construida,casas_decimais=2)
area1 = substituir_ponto_por_virgulas(area1,casas_decimais=4)
area2 = substituir_ponto_por_virgulas(area2,casas_decimais=4)

#convertendo os valores para moeda
valor_mercado_limpo = formatar_valor(valor_mercado)
liquidacao_forcada_limpo = formatar_valor(liquidacao_forcada)

#chamando funções
data_atual = gerar_data_atual()
mercado_extenso = valor_por_extenso(valor_mercado)
liquidacao_extenso = valor_por_extenso(liquidacao_forcada)
declividades = selecionar_declividade()#lista de declividades e suas descrições
pedologias = selecionar_pedologia()#lista de pedologias e suas descrições
texto_das_desclividades = gerar_texto(declividades)#aqui vai gerar um texto com todas as declividades separando elas por paaragrafo
inciais_declividades = extrair_iniciais_desclividades(declividades)#aqui ele vai pegar as inicias das declividades
texo_mosaico = fazer_Texto_mosaico(inciais_declividades)#aqui ele vai gerar um texto que vai pro mosaico
titulo_declividades = fazer_titulo_Declividade(inciais_declividades)#aqui ele vai fazer o titulo da declividade
texto_das_pedologias = gerar_texto(pedologias)#aqui vai gerar um texto com todas as pedologia separando elas por paaragrafo
nomes_pedologias = extrair_nomes_pedologias(pedologias)#aqui eu ele vai pegar apenas os nomes das pedologias da lista de pedologias selecionadas pelo usuario
titulo_pedologias = fazer_texto_pedologia(nomes_pedologias)#aqui ele vai fazer o texto do titulo das pedologias
rota_de_acesso = escolher_e_ler_arquivo_txt()





#tratando se é cpf ou cnpj 
#propietario
if genero == 'empresa' or len(cpf_cpnj_propietario) == 18:
    reposta_cpf_cnpj_propietario = "CNPJ"
else:
    reposta_cpf_cnpj_propietario = "CPF"


#arrumando bacia e sub-bacia
if sub_bacia == 'None':#aqui eu sei que se a sub-bacia for none eu sei que so vai ter barcia
    resposta_sistema_hidrografico = texto_bacia
else:
    resposta_sistema_hidrografico = texto_sub_bacia

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

#aqui eu to fazendo o texto do valor de mercado do item 14
def resposta_texto_mercado():
    return texto_valor_mercado()

def reposta_valores_de_mercado():
    return valores_mercado()

def resposta_valores_liquidacao_forcada():
    return valores_liquidacao_forcada()

def resposta_texto_do_final_liquidacao_forcada():
    return texto_do_final_da_liquidacao_forcada()

#aqui eu to indo selecionar os mapas da minha aplicação
lista_de_imagens_mapas = selecionar_imagens_retornar_caminho('Selecione os MAPAS')
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
croqui_imagem = imagem_croqui()
#aqui eu to inserindo  a imagem de layout geral na capa do laudo
inserir_layout_geral_na_capa(layout_geral_mapa,celula_para_colar_o_mapa)

#aqui eu to copiando a tabela que vai na capa para a tabela de estatistica
destino_excel_estatistica = selecionar_arquivo_excel()
copiar_pagina_excel(destino_excel_estatistica)

images_with_placeholders = [#essa variavel aqui contem uma lista com os caminhos dos mapas
    (pedologia_mapa, "{{h01hf}}"),
    (bacia_ou_sub_bacia_mapa, "{{h9fd1}}"),
    (declividade_mapa, "{{g7aa}}"),
    (croqui_imagem, "{{dh19a}}"),
]

#Insere as imagens substituindo os marcadores de posição
colar_maps_e_croquii(word_filee, images_with_placeholders)

#arrumar cpf ou cnpj do proponente
arrumar_cpf_cnpj_proponente(destino_excel_estatistica)
caminho_imagens_cit = selecionar_imagens_retornar_caminho('SELECIONE TODAS AS CIT')
caminho_imagens_cnd = selecionar_imagens_retornar_caminho('SELECIONE TODAS AS CND')
caminho_imagens_ccir = selecionar_imagens_retornar_caminho('SELECIONE TODOS OS CCIR')
caminho_imagens_car = selecionar_imagens_retornar_caminho('SELECIONE TODOS OS CAR')
images_by_code = { #aqui tem uma lista de com os caminhhos dos documentos
    "{{sh19}}": caminho_imagens_cit,
    "{{s9ah1}}": caminho_imagens_cnd,
    "{{su1h0}}": caminho_imagens_ccir,
    "{{h9d1a}}": caminho_imagens_car,
   
}

# Insere múltiplas imagens por código
colar_imagens_documentos(word_filee, images_by_code)
