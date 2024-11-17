from funcoes import *
from leitura_excel import *
from textos import *




#tratando as datas
data_solicitacao = formatar_data(data_solicitacao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_emissao = formatar_data(data_emissao)#aqui to to tratando a data pra ela ficar assim : 00/00/0000
data_vistoria = formatar_data(data_vistoria)#aqui to to tratando a data pra ela ficar assim : 00/00/0000

#colocando virgulas nas areas ao inves de pontos
area_ha = area_ha.replace(".",",")
area_construida = area_construida.replace(".",",")
area1 = area1.replace(".",",")
area2 = area2.replace(".",",")

#convertendo os valores para moeda
valor_mercado_limpo = formatar_valor(valor_mercado)
liquidacao_forcada_limpo = formatar_valor(liquidacao_forcada)

#chamando funções
data_atual = gerar_data_atual()
mercado_extenso = valor_por_extenso(valor_mercado)
liquidacao_extenso = valor_por_extenso(liquidacao_forcada)

#resposta pro memorial descritivo
#arrumar essa logica aqui
if gel is not None:
    resposta_gel_referenciamento = texto_do_gel
elif gel is None :
    if memorial_descritivo == 'sim':
        resposta_gel_referenciamento = texto_memorial_descritivo
