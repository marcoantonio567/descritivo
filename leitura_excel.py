from funcoes import ler_celula_excel
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")#aqui eu to tirando um aviso de erro que tinha no terminal, para não preucupar os desenvolvedores


proponente = str(ler_celula_excel('c2'))
cpf_cpnj_proponente = str(ler_celula_excel('c3'))
propietario = str(ler_celula_excel('c5'))
cpf_cpnj_propietario = str(ler_celula_excel('c6'))
nome_imovel = str(ler_celula_excel('c8'))
matricula = str(ler_celula_excel('c9'))
identificacao = str(ler_celula_excel('c10'))
area_ha = float(ler_celula_excel('c11'))
area_construida = float(ler_celula_excel('c12'))
valor_mercado = str(ler_celula_excel('c13'))
liquidacao_forcada = str(ler_celula_excel('c14'))
car = str(ler_celula_excel('c15'))
gel = str(ler_celula_excel('c16'))
area1 = float(ler_celula_excel('c17'))
area2 = float(ler_celula_excel('c18'))
municipio = str(ler_celula_excel('c19'))
uf = str(ler_celula_excel('c20'))
agencia = str(ler_celula_excel('f2'))
municipio_agencia = str(ler_celula_excel('f3'))
data_solicitacao = str(ler_celula_excel('f4'))
municipio_cartorio = str(ler_celula_excel('f5'))
data_emissao = str(ler_celula_excel('f6'))
data_vistoria = str(ler_celula_excel('f8'))
fluid = str(ler_celula_excel('c21'))
georeferenciamento = str(ler_celula_excel('c22'))
genero = str(ler_celula_excel('f9'))
hipotecas = str(ler_celula_excel('c23'))
bioma_amazonico = str(ler_celula_excel('c24'))
casamento = str(ler_celula_excel('f10'))
genero_casamento = str(ler_celula_excel('f11'))
nome_casamento = str(ler_celula_excel('f12'))
cpf_casamento = str(ler_celula_excel('f13'))
bacia = str(ler_celula_excel('c25'))
sub_bacia = str(ler_celula_excel('c26'))


# print(f"""
# Proponente: {proponente}s
# CPF/CNPJ do Proponente: {cpf_cpnj_proponente}
# Proprietário: {propietario}
# CPF/CNPJ do Proprietário: {cpf_cpnj_propietario}
# Nome do Imóvel: {nome_imovel}
# Matrícula: {matricula}
# Identificação: {identificacao}
# Área (ha): {area_ha}
# Área Construída: {area_construida}
# Valor de Mercado: {valor_mercado}
# Liquidação Forçada: {liquidacao_forcada}
# CAR: {car}
# GEL: {gel}
# Área 1: {area1}
# Área 2: {area2}
# Município: {municipio}
# UF: {uf}
# Agência: {agencia}
# Município da Agência: {municipio_agencia}
# Data de Solicitação: {data_solicitacao}
# Município do Cartório: {municipio_cartorio}
# Data de Emissão: {data_emissao}
# Data de Vistoria: {data_vistoria}
# georeferenciamento : {georeferenciamento}
# genero: {genero}
# hipotecas : {hipotecas}
# bioma amazonico : {bioma_amazonico}
# """)
