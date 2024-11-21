from funcoes import *
from leitura_excel import *
from textos import *



#tratando as datas
data_solicitacao_limpo = formatar_data(data_solicitacao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_emissao_limpo = formatar_data(data_emissao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_vistoria_limpo = formatar_data(data_vistoria)#aqui to to tratando a data pra ela ficar assim : 00/00/0000

#colocando virgulas nas areas ao inves de pontos
area_ha = substituir_ponto_por_virgulas(area_ha)
area_construida = substituir_ponto_por_virgulas(area_construida)
area1 = substituir_ponto_por_virgulas(area1)
area2 = substituir_ponto_por_virgulas(area2)

#convertendo os valores para moeda
valor_mercado_limpo = formatar_valor(valor_mercado)
liquidacao_forcada_limpo = formatar_valor(liquidacao_forcada)

#chamando funções
data_atual = gerar_data_atual()
mercado_extenso = valor_por_extenso(valor_mercado)
liquidacao_extenso = valor_por_extenso(liquidacao_forcada)
declividades = selecionar_declividade()
pedologias = selecionar_pedologia()
texto_das_desclividades = gerar_texto(declividades)
texto_das_pedologias = gerar_texto(pedologias)
rota_de_acesso = escolher_e_ler_arquivo_txt()

#resposta para o memorial descritivo
if georeferenciamento == 'GEL':
    resposta_gel_referenciamento = texto_do_gel
elif georeferenciamento == 'CAR':
    resposta_gel_referenciamento = texto_car
elif georeferenciamento == 'Memorial descritivo':
    resposta_gel_referenciamento = texto_memorial_descritivo
elif georeferenciamento == 'CAR e Memorial':
    resposta_gel_referenciamento = texto_memorial_descritivo_com_car

#resolvendo questões de tipo de pessoa
if genero == 'empresa':
    resposta_tipo_pessoa = texto_pessoa_juridica
else:
    if casamento == 'sim':
        resposta_tipo_pessoa = texto_pessoa_fisica_casada
    else:
        resposta_tipo_pessoa = texto_pessoa_fisica

#verificando se o imovel tem hipotecas pendentes ou não
if hipotecas == 'sim':
    reposta_hipoteca = texto_imovel_com_hipotecas
elif hipotecas == 'não':
    reposta_hipoteca = texto_imovel_sem_hipotecas

#verificação se o imovel esta inserido no bioma amazonico
if bioma_amazonico == 'sim':
    reposta_bioma = texto_bioma_inserido
elif bioma_amazonico == 'não':
    reposta_bioma = texto_bioma_nao_inserido



#tratando se é cpf ou cnpj 
#propietario
if genero == 'empresa' or len(cpf_cpnj_propietario) == 18:
    reposta_cpf_cnpj_propietario = "CNPJ"
else:
    reposta_cpf_cnpj_propietario = "CPF"

#proponente
if len(cpf_cpnj_proponente) == 14 or len(cpf_cpnj_proponente) == 11:
    reposta_cpf_cnpj_proponente = "CPF"
else:
    reposta_cpf_cnpj_proponente = "CNPJ"


#arrumando bacia e sub-bacia
if sub_bacia == 'None':#aqui eu sei que se a sub-bacia for none eu sei que so vai ter barcia
    resposta_sistema_hidrografico = texto_bacia
else:
    resposta_sistema_hidrografico = texto_sub_bacia
