from funcoes import substituir_palavras_documento , abrir_arquivo_word
from tratamento_dados import *


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
    '#data_solicitacao':data_solicitacao_limpo,
    '#data_vistoria':data_vistoria_limpo,
    '#emissao_cartorio':data_emissao_limpo,
    '#948':municipio_agencia,#agencia do municipio
    '#extensao_mercado':mercado_extenso,
    '#extensao_liquidacao':liquidacao_extenso,
    'texto_georeferencimento':resposta_gel_referenciamento,
    '#texto_tipo_pessoa':resposta_tipo_pessoa,
    '#texto_hipotecas':reposta_hipoteca,
    '#7452':reposta_bioma,#verificando se o imovel ta inserido no bioma mazonioco ou não
    '#486':reposta_cpf_cnpj_proponente,#cpf ou cnpj do proponente
    '#485':reposta_cpf_cnpj_propietario,#cpf ou cnpj do propietario
    '#rota_acess':rota_de_acesso,
    '#texto_bacia':resposta_sistema_hidrografico,
    
}
saida = 'teste.docx'
substituir_palavras_documento('LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx',dados,saida)
abrir_arquivo_word(saida)
