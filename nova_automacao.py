from funcoes import (substituir_palavras_documento , abrir_arquivo_word , 
                     colocar_quantidade_de_paginas_laudo_e_enumeradas ,procurar_arquivo_word, 
                     renomear_arquivo_word,negritar_texto_entre_tags,
                     colocar_tag_texto)
from tratamento_dados import *
dados = {
    '#proponente':proponente,
    '#propietario':propietario,
    '#agencia':agencia,
    '#nome_imovel':lista_nomes_imoveis_separados_por_e,
    '#municipio':municipios[0],
    '#cpf_proponente':cpf_cpnj_proponente,
    '#cpf_propietario':cpf_cpnj_propietario,
    '#uf':estado_uf[0],
    '#matricula':matricula,
    '#identificacao':identificacao,
    '#area_ha':area_ha_limpo,
    '#area_contruida':area_construida_limpo,
    '#MERCADO':valor_mercado_limpo,
    '#LIQUIDACAO':liquidacao_forcada_limpo,
    '#gel':gel,
    '#CAR':car,
    '#area1':area1_limpo,
    '#area2':area2_limpo,
    '#fluid':fluid,
    '#data_atual':data_atual,
    '#data_solicitacao':colocar_tag_texto(data_solicitacao_limpo),
    '#data_vistoria':colocar_tag_texto(data_vistoria_limpo),
    '#emissao_cartorio':colocar_tag_texto(data_emissao_limpo),
    '#948':colocar_tag_texto(municipio_agencia),#agencia do municipio
    '#juki12':fazer_texto_georeferenciamento(),
    '#texto_tipo_pessoa':texto_tipo_pessoa(),
    '#texto_hipotecas':hipotecas_reposta(),
    '#7452':reposta_bioma(),#verificando se o imovel ta inserido no bioma mazonioco ou não
    '#485':reposta_cpf_cnpj_propietario,#cpf ou cnpj do propietario
    '#486':reposta_cpf_cnpj_proponente,#cpf ou cnpj do proponente
    '#rota_acess':rota_de_acesso,
    '#texto_bacia':resposta_sistema_hidrografico(),
    '#3498':texto_das_desclividades,
    '#3497':texto_das_pedologias,
    '#5271':registro_imovel(),
    '#huas123':titulo_pedologias,
    '#ua5yi1':texo_mosaico,
    'g0jd1':resposta_introducao_declividade,
    '#dj10f':resposta_texto_mercado(),
    'h8f1a':reposta_valores_de_mercado(),
    'Gfas8':resposta_valores_liquidacao_forcada(),
    '#jd01a':resposta_texto_do_final_liquidacao_forcada(),
    '9ashd':texto_coeficientes,
    'jsad0':declividade_vegetação,
    'as8ghd9a':nomes_imoveis_formatado,
    'As0k':resposta_texto_data_emissao(),
    'biasd91':respota_da_ia().replace('*',''),
    '#8g7a8':resposta_texto_descrição_imovel_avaliando(),
    
    
    
}

print(f'{DARK_GREEN}carregando mapa do layouyt gelrap {layout_geral_mapa}{RESET}')
caminho_antigo = f'teste.docx'
substituir_palavras_documento(caminho_antigo,dados,caminho_antigo)
colocar_quantidade_de_paginas_laudo_e_enumeradas(quantidade_de_imagens)
nome_atualizado = renomear_arquivo_word(caminho_antigo,texto_nome_arquivo)#tive que criar uma variavel porque como o final é um nome dinamico a respeito da quantidade e do contador de paginas ele so renomeia o arquivo depois fica mais facil doque reconstruir a função
negritar_texto_entre_tags(nome_atualizado)
copiar_ou_recortar_arquivos(r'C:\\Users\\USUARIO\\Desktop\\automacao_descritivo_laudo\\'
                            ,pasta_pecas_tecnicas, 'recortar')
copiar_ou_recortar_arquivos(r'C:\\Users\\USUARIO\\Desktop\\automacao_descritivo_laudo\\TEMPLATES',
                            r'C:\\Users\\USUARIO\\Desktop\\automacao_descritivo_laudo\\', 'copiar')
laudo_completo = procurar_arquivo_word(pasta_pecas_tecnicas)
abrir_arquivo_word(laudo_completo)
