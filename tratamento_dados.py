from funcoes import *
from leitura_excel import *
from textos import *

word_filee = 'LAUDO DE AVALIAÇÃO PARA AUTOMAÇÃO.docx'
substituir_cabecalho(texto_cabecalho,word_filee,word_filee)
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

#resposta para o memorial descritivo
def fazer_texto_georeferenciamento():
    resposta_gel_referenciamento = []
    for item in data_imoveis:
        if georeferenciamento == 'GEL':
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



