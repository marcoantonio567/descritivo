from leitura_excel import *
from fazer_capa_excel import contar_dados_vazios, ler_matriculas ,ler_propietarios , ler_casamento  
from funcoes import valor_por_extenso , separar_ultimo_item_por_e , agrupar_Matricula_por_Car , fazer_Texto_atraves_De_lista , agrupar_Matricula_por_data_Emissao , formatar_data
from cores import *

quantidade_matriculas = 10 - contar_dados_vazios(coluna='G')
quantidade_propietarios = 10 - contar_dados_vazios()
data_imoveis = ler_matriculas(quantidade_matriculas)#dados dos imoveis SAIDA: matricula,identificacao,area_ha,area_contruida,valor_mercado,liquidacao_forcada
print(f'{DARK_RED}{data_imoveis}{RESET}')
dados_propietarios = ler_propietarios(quantidade_propietarios)
casamento_propietarios = ler_casamento(quantidade_propietarios)
apenas_matriculas = [str(dado[0]) for dado in data_imoveis]#aqui e pra ele extrair apenas as matriculas dos imoveis
apenas_nomes_imoveis = [str(dado[6]) for dado in data_imoveis]#aqui e pra ele extrair apenas os nomes dos imoveis
lista_matriculas_separadas_por_e = " e ".join(apenas_matriculas)
lista_nomes_imoveis_separados_por_e = " e ".join(apenas_nomes_imoveis)
nomes_imoveis_formatado = separar_ultimo_item_por_e(apenas_nomes_imoveis)#essa função pega uma lista e troca a separação do ultimo item de virgula pra "e"
lista_matriculas_formatada = separar_ultimo_item_por_e(apenas_matriculas)


texto_cabecalho = f'LAUDO DE AVALIAÇÃO N° {fluid} - {proponente.upper()} - {agencia.upper()}'#aqui to mostarando qual vai ser o texto do cabeçalho
texto_nome_arquivo = f'LAUDO DE AVALIAÇÃO Nº {fluid} - {proponente}'


def texto_vegetacao_titulos(vegetacao,quatitade_imoveis,ponta_titulo):
    if len(vegetacao) == 1:
        if quatitade_imoveis ==1:
            return f'A vegetação do imóvel é caracterizada pela região fitoecológica classificada como <tag>{ponta_titulo}</tag>'
        else:
            return f'A vegetação dos imóveis é caracterizada pela região fitoecológica classificada como <tag>{ponta_titulo}</tag>'
    else:
        if quatitade_imoveis ==1:
            return f'A vegetação do imóvel é caracterizada pela região fitoecológica classificadas como <tag>{ponta_titulo}</tag>'
        else:
            return f'A vegetação dos imóveis é caracterizada pela região fitoecológica classificada como <tag>{ponta_titulo}</tag>'
def texto_declividades_titulo(declividades,quatitade_imoveis,ponta_titulo):
    if len(declividades) == 1:
        if quatitade_imoveis ==1:
            return f'O relevo do imóvel possui ondulações. Sendo caracterizado pela declividade classificada como <tag>{ponta_titulo}</tag>'
        else:
            return f'O relevo dos imóvel possui ondulações. Sendo caracterizado pela declividade classificada como <tag>{ponta_titulo}</tag>'
    else:
        if quatitade_imoveis ==1:
            return f'O relevo do imóvel possui ondulações. Sendo caracterizado pelas declividades classificadas como <tag>{ponta_titulo}</tag>'
        else:
            return f'O relevo dos imóvel possui ondulações. Sendo caracterizado pela declividade classificada como <tag>{ponta_titulo}</tag>'
def texto_memorial_descritivo(id_fazenda): 
    return f'O imóvel de matrícula <tag>{id_fazenda}</tag> não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo'
def texto_memorial_descritivo_com_car(id_fazenda): 
    return f'O imóvel de matrícula <tag>{id_fazenda}</tag> não possui georreferenciamento averbado em matrícula, sua localização foi obtida através de coordenadas do memorial descritivo e do Cadastro Ambiental Rural - CAR.'
def texto_do_gel(id_fazenda): 
    return f'O imóvel de matrícula n° <tag>{id_fazenda}</tag>  possui georreferenciamento averbado em matrícula, certificado pelo Instituto Nacional de Colonização e Reforma Agrária - INCRA, consoante a certificação nº <tag>{gel}</tag>. Sua localização foi obtida através do mesmo.\n'
def texto_car(id_fazenda): 
    return f'O imóvel de matrícula nº <tag>{id_fazenda}</tag> não possui georeferenciamento averbado em matrícula. Sua localização foi obtida através do Cadastro Ambiental Rural.'
def texot_car_geo(id_fazenda):
    return f'O imóvel de matrícula n° <tag>{id_fazenda}</tag>  possui georreferenciamento averbado em matrícula, certificado pelo Instituto Nacional de Colonização e Reforma Agrária - INCRA, consoante a certificação nº <tag>{gel}</tag>. e do Cadastro Ambiental Rural.'
def texto_tipo_pessoa():
    quantidade = quantidade_propietarios
    #propietarios:
    #nome,cpf_cnpj,porcentagem,casado,genero

    #casamentos
    #genero,nome,cpf
    if quantidade>1:
        texto_cabecalho = 'Os imóveis em questão são de propriedade '
    else:
        texto_cabecalho = 'O imóvel em questão é de propriedade '
    escopo_Texto = []
    for dado ,dado2 in zip(dados_propietarios,casamento_propietarios):
        
        if dado[3] == 'Sr.':#aqui eu to verificando se ele e homen ou mulher
            if dado[4] == 'sim':#aqui to veficando se ele é casado
                if str(0)  in str(dado[2]):#aqui eu to verificando se o cpf não e vazio
                    texto = f'do <tag>Sr. {dado[0]}</tag> inscrito no CPF nº <tag>{dado[1]}</tag>, casado com a <tag>Sra. {dado2[1]}</tag>, inscrita no CPF nº <tag>{dado2[2]}</tag>.'
                    escopo_Texto.append(texto)
                else:
                    texto = f'do <tag>Sr. {dado[0]}</tag> inscrito no CPF nº <tag>{dado[1]}</tag>, casado com a <tag>Sra. {dado2[1]}</tag>, inscrita no CPF nº <tag>{dado2[2]}</tag>, que detêm <tag>{dado[2]}%</tag> do imovel.'
                    escopo_Texto.append(texto)
            else:
                texto = f'do <tag>Sr. {dado[0]}</tag> inscrito no CPF nº <tag>{dado[1]}</tag>, que detêm <tag>{dado[2]}%</tag> do imovel.'
                escopo_Texto.append(texto)
        elif dado[3] == 'Sra.':
            if dado[4] == 'sim':
                if str(0)  in str(dado[2]):#aqui eu to verificando se o cpf não e vazio
                    texto = f'da <tag>Sra. {dado[0]}</tag> inscrita no CPF nº <tag>{dado[1]}</tag>, casada com o <tag>Sr. {dado2[1]}</tag>, inscrito no CPF nº <tag>{dado2[2]}</tag>.'
                    escopo_Texto.append(texto)
                else:
                    texto = f'da <tag>Sra. {dado[0]}</tag> inscrita no CPF nº <tag>{dado[1]}</tag>, casada com o <tag>Sr. {dado2[1]}</tag>, inscrito no CPF nº <tag>{dado2[2]}</tag>, que detêm <tag>{dado[2]}%</tag> do imveo.'
                    escopo_Texto.append(texto)
            else:
                texto = f'da <tag>Sra. {dado[0]}</tag> inscrita no CPF nº <tag>{dado[1]}</tag>, que detêm <tag>{dado[2]}%</tag> do imovel.'
                escopo_Texto.append(texto)
        else:
            texto = f'da empresa <tag>{dado[0]}</tag>, inscrita no cnpj <tag>{dado[1]}</tag>.'
            escopo_Texto.append(texto)

    texto_output = " ".join(escopo_Texto)
    return texto_cabecalho+texto_output
def texto_registro(quantidade):
    texto_atualizado = agrupar_Matricula_por_Car(data_imoveis)
    if texto_atualizado == 'todos itens sao iguais':
        if quantidade <=1:
            return f'O imóvel está registrado no CAR: <tag>{car}</tag>. Foi detectada diferença entre a área do imóvel rural declarada conforme documentação comprobatória de propriedade [<tag>{area1}</tag> hectares] e a área do imóvel rural identificada em representação gráfica [<tag>{area2}</tag> hectares].'
        else:
            return f'Os imóveis estão registrados no CAR: <tag>{car}</tag>. Foi detectada diferença entre a área do imóvel rural declarada conforme documentação comprobatória de propriedade [<tag>{area1}</tag> hectares] e a área do imóvel rural identificada em representação gráfica [<tag>{area2}</tag> hectares].'
    else:
        texto_definitivo = fazer_Texto_atraves_De_lista(texto_atualizado)
        return texto_definitivo
def texto_hipotecas(quantidade):
    if quantidade <=1:
        if hipotecas == 'sim':
            return  f'De acordo com o registro da matrícula n° <tag>{matricula}</tag>, o imóvel possui registro de hipoteca pendente.'
        else:
            return  f'Conforme averbações da matrícula n° <tag>{matricula}</tag>, o imóvel não possui hipotecas pendentes.'
            
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
def texto_valor_mercado(municipio,uf):
 
    #aqui eu to seprarando as matriculas por "e" ao inves de virgula
   
    #aqui eu to fazendo uma estrutura de condição pra ver qual o texto que vou usar
    if quantidade_matriculas >1:
        texto = f'Com base em todo o exposto apresentado, diante da metodologia de trabalho bem como as planilhas de cálculo, apurou-se que aos imóveis denominados <tag>{nomes_imoveis_formatado}</tag>, objeto das matrículas <tag>{lista_matriculas_formatada}</tag>, localizados na zona rural do município de {municipio} - {uf}, os seguintes valores:'
    else:
        texto = f'Com base em todo o exposto apresentado, diante da metodologia de trabalho bem como as planilhas de cálculo, apurou-se que o imóvel rural denominado <tag>{nome_imovel}</tag>, localizada no município de <tag>{municipio}</tag> - <tag>{uf}</tag>, o seguinte valor:'
        
    return texto
def valores_mercado(lista_do_data_imoveis):
    #essa função aqui vai basicamente colocar o valor em real é o valor em extenso entre parenteses depos
    texto_unificado = []
    for dado in lista_do_data_imoveis:
        texto_jogar_dentro = f'MATRÍCULA – {dado[0]} - {dado[4]} ({valor_por_extenso(dado[4])})'
        texto_unificado.append(texto_jogar_dentro)
    
    # Junta todos os textos em um único texto
    texto_final = "\n".join(texto_unificado)

    return texto_final
def valores_liquidacao_forcada(lista_do_data_imoveis):
    texto_unificado = []
    for dado in lista_do_data_imoveis:
        texto_jogar_dentro = f'MATRÍCULA – {dado[0]} - {dado[5]} ({valor_por_extenso(dado[5])})'
        texto_unificado.append(texto_jogar_dentro)
    
    # Junta todos os textos em um único texto
    texto_final = "\n".join(texto_unificado)

    return texto_final
def texto_do_final_da_liquidacao_forcada():
    if quantidade_matriculas >1:
        texto = 'os seguintes valores'
    else:
        texto = 'o seguinte valor'
        
    return texto
def sistema_hidrografico():
    if sub_bacia == 'None':#aqui eu sei que se a sub-bacia for none eu sei que so vai ter barcia
        if quantidade_matriculas >1:
            return f'Os imóveis estão contidos dentro do Sistema Hidrográfico Bacia do <tag>{bacia}</tag>.'
        else:
            return f'O imóvel está contido dentro do Sistema Hidrográfico Bacia do <tag>{bacia}</tag>.'
    else:
        if quantidade_matriculas>1:
            return f'Os imóveis estão contido dentro do Sistema Hidrográfico Bacia do <tag>{bacia}</tag>, e Sub-Bacia <tag>{sub_bacia}</tag>.'
        else:
            return f'O imóvel está contido dentro do Sistema Hidrográfico Bacia do <tag>{bacia}</tag>, e Sub-Bacia <tag>{sub_bacia}</tag>.'
def texto_data_emissao():
    texto_atualizado = agrupar_Matricula_por_data_Emissao(data_imoveis,municipio_agencia)
    print(f'{DARK_RED}{texto_atualizado}{RESET}')
    if texto_atualizado == 'todos itens sao iguais':
        return f'Matrícula(s) n° <tag>{lista_matriculas_formatada}</tag> do Cartório de {municipio_agencia}, emitida no dia <tag>{formatar_data(data_emissao)}</tag>.'
    else:
        texto_definitivo = fazer_Texto_atraves_De_lista(texto_atualizado)
        return texto_definitivo	
    
def texto_descrição_imovel_avaliando(lista_municipios):
    if len(lista_matriculas_separadas_por_e) >1:
        #lista_matriculas_separadas_por_e
        return f'As matrículas {lista_matriculas_separadas_por_e}, localizadas no município de {lista_municipios[0]}, compõe {lista_nomes_imoveis_separados_por_e}, voltada para atividade pecuária. Uma descrição mais detalhada das matrículas pode ser conferida nos quadros a seguir:'
    else:
        return f'A matrícula {lista_matriculas_separadas_por_e}, localizada no município de {lista_municipios[0]}, compõe {lista_nomes_imoveis_separados_por_e}, voltada para atividade pecuária. Uma descrição mais detalhada das matrícula pode ser conferida no quadro a seguir:'