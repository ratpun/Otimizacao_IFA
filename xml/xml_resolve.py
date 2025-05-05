import xmltodict

with open("C:/Users/rafaellazaro/Downloads/XML E PDF FARMANIA 042024 À 032025/XML/Notas_384946025-04-2025_Chave_42240409026759000118550010002634381183340120S_57.xml", "rb") as arquivo_insumos:
    dic_notaFiscal = xmltodict.parse(arquivo_insumos.read())

insumos = dic_notaFiscal['nfeProc']['NFe']['infNFe']['det']
distribuidora = dic_notaFiscal['nfeProc']['NFe']['infNFe']['emit']

empresa = distribuidora['xNome'] # Nome da empresa
cnpj = distribuidora['CNPJ'] # CNPJ da empresa

for prod in insumos:
    codigo = prod['prod']['cProd'] # Código do produto
    descricaoExtendida = prod['prod']['xProd'] # Descrição extendida do produto
    descricao = descricaoExtendida[0:descricaoExtendida.index('--')].strip() # Descrição do produto
    gramaturaVendida = descricaoExtendida[descricaoExtendida.index('--')+2:descricaoExtendida.index('Lote')].strip() # Gramatura vendida do produto
    ncm = prod['prod']['NCM'] # NCM do produto
    unidade = prod['prod']['uCom'] # Unidade de medida do produto
    quantidade = prod['prod']['qCom'] # Quantidade do produto
    valorUnit = prod['prod']['vUnCom'] # Valor unitário do produto
    valorTotal = prod['prod']['vProd'] # Valor total do produto
    valorFrete = prod['prod']['vFrete'] # Valor do frete do produto
    dataFab = prod['prod']['rastro']['dFab'] # Data de fabricação do produto
    dataVal = prod['prod']['rastro']['dVal'] # Data de validade do produto
    valorICMS = prod['imposto']['ICMS']['ICMS00']['vICMS'] # Valor do ICMS do produto
    valorBaseCalcICMS = prod['imposto']['ICMS']['ICMS00']['vBC'] # Valor da base de cálculo do ICMS do produto
    aliquotaICMS = prod['imposto']['ICMS']['ICMS00']['pICMS'] # Alíquota do ICMS do produto
    produtoCompleto = [codigo, descricao, gramaturaVendida, ncm, unidade, quantidade, valorUnit, valorTotal, valorFrete, dataFab, dataVal, valorICMS, valorBaseCalcICMS, aliquotaICMS]
    print(produtoCompleto)
