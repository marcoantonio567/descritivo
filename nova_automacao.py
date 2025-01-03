from funcoes import (substituir_palavras_documento , abrir_arquivo_word , 
                     colocar_quantidade_de_paginas_laudo ,renovar_a_integração , 
                     renomear_arquivo_word,negritar_texto_entre_tags)
from tratamento_dados import *


#rodar_extracao_tabelas()#aqui to extraindo todas as tabelas primeiro

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
    '#1972':mercado_extenso,
    '1985':liquidacao_extenso,
    '#juki12':fazer_texto_georeferenciamento(),
    '#texto_tipo_pessoa':texto_tipo_pessoa(),
    '#texto_hipotecas':hipotecas_reposta(),
    '#7452':reposta_bioma(),#verificando se o imovel ta inserido no bioma mazonioco ou não
    '#485':reposta_cpf_cnpj_propietario,#cpf ou cnpj do propietario
    '#rota_acess':rota_de_acesso,
    '#texto_bacia':resposta_sistema_hidrografico,
    '#3498':texto_das_desclividades,
    '#3497':texto_das_pedologias,
    '#5271':registro_imovel(),
    '#huas123':titulo_pedologias,
    '#ua5yi1':texo_mosaico,
    'g0jd1':titulo_declividades,
    '#dj10f':resposta_texto_mercado(),
    'h8f1a':reposta_valores_de_mercado(),
    'Gfas8':resposta_valores_liquidacao_forcada(),
    '#jd01a':resposta_texto_do_final_liquidacao_forcada(),
    
    
    
}

caminho_antigo = f'teste.docx'
substituir_palavras_documento('LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx',dados,caminho_antigo)
colocar_quantidade_de_paginas_laudo()
renovar_a_integração()#aqui eu to copiando o excel template para a aplicação real(deixar inicialmente desabilitado vai depender de quem esta usando)
caminho_novo = renomear_arquivo_word(caminho_antigo,texto_nome_arquivo)#tive que criar uma variavel porque como o final é um nome dinamico a respeito da quantidade e do contador de paginas ele so renomeia o arquivo depois fica mais facil doque reconstruir a função
negritar_texto_entre_tags(caminho_novo)
abrir_arquivo_word(caminho_novo)
