import xmltodict
import os
from openpyxl import Workbook

# Este script percorre arquivos XML de notas fiscais, extrai informações relevantes e as organiza em um arquivo Excel.
# As informações extraídas incluem nome da empresa, data de emissão, código do produto, descrição, NCM, unidade de medida, quantidade, valor unitário, valor total, data de fabricação e data de validade.

def cria_resumo_excel():
    pasta_mae = "C:/Users/rafaellazaro/Downloads/XML E PDF FARMANIA 042024 À 032025/XML"
    pasta_excel = "C:/Users/rafaellazaro/Downloads/XML E PDF FARMANIA 042024 À 032025"
    total_comprado = percorre_arquivos(pasta_mae) # Percorre os arquivos na pasta
    cria_excel(total_comprado, pasta_excel) # Cria o arquivo Excel com os dados extraídos
    

def extrai_dados_nota_fiscal(arquivo):
    lista_produtos_comprados = [] # Lista de produtos comprados
    with open(arquivo, encoding='utf8', errors='ignore') as arquivo_insumos:
        dic_notaFiscal = xmltodict.parse(arquivo_insumos.read())

    insumos = dic_notaFiscal['nfeProc']['NFe']['infNFe']['det']
    dataEmissao = dic_notaFiscal['nfeProc']['NFe']['infNFe']['ide']['dhEmi'] # Data de emissão da nota fiscal
    dataEmissao = dataEmissao.split('T')[0] # Remove a hora da data de emissão
    distribuidora = dic_notaFiscal['nfeProc']['NFe']['infNFe']['emit']

    nomeEmpresa = distribuidora['xNome'] # Nome da empresa
    cnpj = distribuidora['CNPJ'] # CNPJ da empresa
    empresa = nomeEmpresa + ' - ' + cnpj # Nome completo da empresa


    if isinstance(insumos, dict):  # Ensure insumos is a list
        insumos = [insumos]

    for prod in insumos:
        codigo = prod['prod']['cProd'] # Código do produto
        descricaoExtendida = prod['prod']['xProd'] # Descrição extendida do produto
        ncm = prod['prod']['NCM'] # NCM do produto
        unidade = prod['prod']['uCom'] # Unidade de medida do produto
        quantidade = prod['prod']['qCom'] # Quantidade do produto
        valorUnit = prod['prod']['vUnCom'] # Valor unitário do produto
        valorTotal = prod['prod']['vProd'] # Valor total do produto
        dataFab = None
        dataVal = None
        if 'rastro' in prod['prod']:
            if 'dFab' in prod['prod']['rastro']:
                dataFab = prod['prod']['rastro']['dFab'] # Data de fabricação do produto
                dataVal = prod['prod']['rastro']['dVal'] # Data de validade do produto
        produtoCompleto = [empresa, dataEmissao, codigo, descricaoExtendida, ncm, unidade, quantidade, valorUnit, valorTotal, dataFab, dataVal]
        lista_produtos_comprados.append(produtoCompleto) # Adiciona o produto à lista de compras
    return lista_produtos_comprados # Retorna a lista de produtos comprados

def percorre_arquivos(pasta_mae):
    arquivos = os.listdir(pasta_mae) # Lista de arquivos na pasta
    produtos_por_empresa = []
    for arquivo in arquivos:
        caminho_completo = os.path.join(pasta_mae, arquivo) # Caminho completo do arquivo
        if os.path.isfile(caminho_completo):  # Verifica se é um arquivo
            print(f"Processando arquivo: {caminho_completo}")
            produtos_por_empresa.append(extrai_dados_nota_fiscal(caminho_completo)) # Extrai os dados da nota fiscal
    return produtos_por_empresa # Retorna a lista de produtos por empresa

def cria_excel(total_comprado, pasta_excel):
    wb = Workbook() # Cria um novo arquivo Excel
    ws = wb.active # Seleciona a planilha ativa
    ws.title = "Produtos totais" # Define o título da planilha
    ws.append(['Nome da empresa', 'Data de emissão', 'Código do produto', 'Descrição do produto', 'NCM', 'Unidade de medida', 'Quantidade', 'Valor unitário', 'Valor total', 'Data de fabricação', 'Data de validade']) 
    for empresa_produtos in total_comprado:
        for produto in empresa_produtos:
            ws.append(produto)
    wb.save(os.path.join(pasta_excel, 'resumo.xlsx')) # Salva o arquivo Excel com o nome 'resumo.xlsx'
    wb.close() # Fecha o arquivo Excel

cria_resumo_excel() # Chama a função para criar o resumo Excel