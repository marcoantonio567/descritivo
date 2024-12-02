from leitura_excel import *
from fazer_capa_excel import contar_dados_vazios, ler_matriculas


quantidade_matriculas = 10 - contar_dados_vazios(coluna='G')
data_imoveis = ler_matriculas(quantidade_matriculas)
def texto_memorial_descritivo(id_fazenda): return f'O imóvel de matrícula {id_fazenda} não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo'
def texto_memorial_descritivo_com_car(id_fazenda): return f'O imóvel de matrícula {id_fazenda} não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo e do Cadastro Ambiental Rural - CAR.'
def texto_do_gel(id_fazenda): return f'O imóvel de matrícula n° {id_fazenda}  possui georreferenciamento averbado em matrícula, certificado pelo Instituto Nacional de Colonização e Reforma Agrária - INCRA, consoante a certificação nº {gel}. Sua localização foi obtida através do mesmo.'
def texto_car (id_fazenda): return f'O imóvel de matrícula nº {id_fazenda} não possui georeferenciamento averbado em matrícula. Sua localização foi obtida através do Cadastro Ambiental Rural.'

texto_pessoa_juridica = f'Conforme a matrícula n° {matricula}, o imóvel tem como proprietária a pessoa jurídica {propietario}, inscrita no CNPJ de n° {cpf_cpnj_propietario}.'
texto_pessoa_fisica = f'Conforme a matrícula n° {matricula}, o imóvel em questão é de propriedade do {genero} {propietario}, inscrito no CPF nº {cpf_cpnj_propietario}.'
if genero_casamento == 'Sra.':
    texto_pessoa_fisica_casada = f'Conforme a matrícula n° {matricula}, o imóvel tem como proprietário o {genero} {propietario}, inscrito no CPF nº {cpf_cpnj_propietario}, juntamente com sua esposa, a Sra. {nome_casamento}, inscrita no CPF n° {cpf_casamento}.'
elif genero_casamento == 'Sr.':
    texto_pessoa_fisica_casada = f'Conforme a matrícula n° {matricula}, o imóvel tem como proprietário o {genero} {propietario}, inscrito no CPF nº {cpf_cpnj_propietario}, juntamente com seu esposo, o Sr. {nome_casamento}, inscrito no CPF n° {cpf_casamento}.'


def texto_registro(quantidade):
    if quantidade <=1:
        return f'O imóvel está registrado no CAR: {car}. Foi detectada diferença entre a área do imóvel rural declarada conforme documentação comprobatória de propriedade [{area1} hectares] e a área do imóvel rural identificada em representação gráfica [{area2} hectares].'
    else:
        return f'Os imóveis estão registrados no CAR: {car}. Foi detectada diferença entre a área do imóvel rural declarada conforme documentação comprobatória de propriedade [{area1} hectares] e a área do imóvel rural identificada em representação gráfica [{area2} hectares].'

def texto_hipotecas(quantidade):

    if quantidade <=1:
        if hipotecas == 'sim':
            return  f'De acordo com o registro da matrícula n° {matricula}, o imóvel possui registro de hipoteca pendente.'
        else:
            return  f'Conforme averbações da matrícula n° {matricula}, o imóvel não possui hipotecas pendentes.'
            
    elif quantidade >1:
        if hipotecas == 'sim':
            return  f'Os imóveis possuem Alienação Fiduciária averbada em matrícula'
        else:
            return  f'Os imóveis não possuem hipotecas pendentes.'

def texto_bioma(quantidade):
    if quantidade <=1:
        if bioma_amazonico == 'sim':
            return f'O imóvel está inserido no Bioma Amazônico.'
        else:
            return f'O imóvel não está inserido no Bioma Amazônico.'
    elif quantidade >1:
        if bioma_amazonico == 'sim':
            return f'Os imóveis estão inseridos no Bioma Amazônico.'
        else:
            return f'Os imóveis não estão inseridos no Bioma Amazônico.'
    
texto_bacia = f'O imóvel está contido dentro do Sistema Hidrográfico Bacia do {bacia}.'
texto_sub_bacia = f'O imóvel está contido dentro do Sistema Hidrográfico Bacia do {bacia}, e Sub-Bacia {sub_bacia}.'

