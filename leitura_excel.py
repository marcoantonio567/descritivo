from funcoes import ler_celula_excel
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")#aqui eu to tirando um aviso de erro que tinha no terminal, para não preucupar os desenvolvedores



proponente = str(ler_celula_excel('c2'))
cpf_cpnj_proponente = str(ler_celula_excel('c3'))
propietario = str(ler_celula_excel('c4'))
cpf_cpnj_propietario = str(ler_celula_excel('c5'))
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
agencia = str(ler_celula_excel('f2'))
municipio_agencia = str(ler_celula_excel('f3'))
data_solicitacao = str(ler_celula_excel('f4'))
data_emissao = str(ler_celula_excel('f6'))
data_vistoria = str(ler_celula_excel('f8'))
fluid = str(ler_celula_excel('c21'))
georeferenciamento = str(ler_celula_excel('c22'))
genero = str(ler_celula_excel('f9'))
hipotecas = str(ler_celula_excel('c23'))
bioma_amazonico = str(ler_celula_excel('c24'))
casamento = str(ler_celula_excel('f10'))
bacia = str(ler_celula_excel('c25'))
sub_bacia = str(ler_celula_excel('c26'))

