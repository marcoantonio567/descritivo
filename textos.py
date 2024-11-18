from leitura_excel import *

texto_memorial_descritivo = f'O imóvel de matrícula {matricula} não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo'
texto_memorial_descritivo_com_car = f'O imóvel de matrícula {matricula} não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo e do Cadastro Ambiental Rural - CAR.'
texto_do_gel = f'O imóvel de matrícula n° {matricula}  possui georreferenciamento averbado em matrícula, certificado pelo Instituto Nacional de Colonização e Reforma Agrária - INCRA, consoante a certificação nº {gel}. Sua localização foi obtida através do mesmo.'
texto_car = f'O imóvel de matrícula nº {matricula} não possui georeferenciamento averbado em matrícula. Sua localização foi obtida através do Cadastro Ambiental Rural.'

texto_pessoa_juridica = f'Conforme a matrícula n° {matricula}, o imóvel tem como proprietária a pessoa jurídica {propietario}, inscrita no CNPJ de n° {cpf_cpnj_propietario}.'
texto_pessoa_fisica = f'Conforme a matrícula n° {matricula}, o imóvel em questão é de propriedade do {genero} {propietario}, inscrito no CPF nº {cpf_cpnj_propietario}.'


texto_imovel_sem_hipotecas = f'Conforme averbações da matrícula n° {matricula}, o imóvel não possui hipotecas pendentes.'
texto_imovel_com_hipotecas = f'De acordo com o registro da matrícula n° {matricula}, o imóvel possui registro de hipoteca pendente.'

texto_bioma_inserido = f'O imóvel está inserido no Bioma Amazônico.'
texto_bioma_nao_inserido = f'O imóvel não está inserido no Bioma Amazônico.'

