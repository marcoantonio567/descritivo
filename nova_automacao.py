from funcoes import *
from tratamento_dados import *
from textos import *


dados = {
    '#proponente':proponente,
    '#propietario':propietario,
    '#agencia':agencia,
    '#nome_imovel':nome_imovel,
    '#municipio':municipio,
    '#cpf_proponente':cpf_cpnj_proponente,
    '#cpf_propietario':cpf_cpnj_propietario,
    '#uf':uf,
    '#matricula':matricula,
    '#identificacao':identificacao,
    '#area_ha':area_ha,
    '#area_contruida':area_construida,
    '#MERCADO':valor_mercado_limpo,
    '#LIQUIDACAO':liquidacao_forcada_limpo,
    '#gel':gel,
    '#CAR':car,
    '#area1':area1,
    '#area2':area2,
    '#fluid':fluid,
    '#data_atual':data_atual,
    '#data_solicitacao':data_solicitacao,
    '#municipio_agencia':municipio_agencia,
    '#data_vistoria':data_vistoria,
    '#extensao_mercado':mercado_extenso,
    '#extensao_liquidacao':liquidacao_extenso,
    'texto_georeferencimento':resposta_gel_referenciamento,

}
saida = 'teste.docx'
substituir_palavras_documento('LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx',dados,saida)
abrir_arquivo_word(saida)
