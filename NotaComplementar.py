import os
from bs4 import BeautifulSoup as Bs
import pandas as pd

periodo = int(input('Informe o período: '))

planilha = pd.read_excel(rf"CMED{periodo}.xlsx", sheet_name='Planilha1')
dictcmed = {}
for linha in planilha.iterrows():
    try:
        dictcmed[str(int(linha[1][1]))] = linha[1][0], linha[1][2], linha[1][3], linha[1][4], linha[1][5]
        print(int(linha[1][1]))
    except ValueError:
        pass

reducaodict = {'NEGATIVA': (26.55, 41.36, 25.56), 'POSITIVA': (11.98, 27.88, 16.16), 'NEUTRA': (7.71, 0, 7.10)}
diretorio = 'C:/Users/Paulo Lira/Desktop/Loja 26/Modelo 55/NFes/'
listagemarquivos = os.listdir(diretorio)
cnpj = '27849963000110'
listaca = []
count = 0
aliquotatxt = open('Base_Aliquota.txt', 'r')
aliquotadict = {}
for linha in aliquotatxt:
    aliquotadict[linha.split()[0]] = linha.split()[1]
listaca.append('Chave|Nº Item|EAN|Reg. Anvisa|CódigoProd|NomeProd|NCM|Quantidade|'
               'Valor Unitário|Valor Total|Base ST Retido XML|ICMS ST XML|Alíquota|PMC|PMC 12|'
               'PMC 18|Concessão|Tipo Produto|Fator Redução\n')
for xml in listagemarquivos:
    if xml[6:20] == cnpj and int(xml[2:6]) == periodo:
        with open(diretorio + xml, 'r') as notafiscal:
            dados = notafiscal.readlines()
            dados = "".join(dados)
            bs_dados = Bs(dados, 'xml')
            nitem = bs_dados.find_all('det')
            for pos, item in enumerate(nitem):
                bs_dados2 = Bs(str(item), 'xml')
                ean = bs_dados2.find('cEANTrib').text
                cprod = bs_dados2.find('cProd').text
                descricao = bs_dados2.find('xProd').text
                ncm = bs_dados2.find('NCM').text
                qTrib = float(bs_dados2.find('qTrib').text)
                vUnTrib = float(bs_dados2.find('vUnTrib').text)
                vProd = bs_dados2.find('vProd').text
                vBCSTRet = bs_dados2.find('vBCSTRet')
                vPMC = bs_dados2.find('vPMC')
                vICMSSTRet = bs_dados2.find('vICMSSTRet')
                try:
                    aliquota = aliquotadict[cprod]
                except KeyError:
                    aliquota = ''
                try:
                    reganvisa = dictcmed[ean][0]
                    tipoproduto = dictcmed[ean][1]
                    pmc12 = dictcmed[ean][2]
                    pmc18 = dictcmed[ean][3]
                    lcct = dictcmed[ean][4]
                    if tipoproduto == 'Referência':
                        fatorreducao = reducaodict[lcct][0]
                    elif tipoproduto == 'Genérico':
                        fatorreducao = reducaodict[lcct][1]
                    elif tipoproduto == 'Similar':
                        fatorreducao = reducaodict[lcct][2]
                except KeyError:
                    reganvisa = ''
                    tipoproduto = ''
                    pmc12 = ''
                    pmc18 = ''
                    lcct = ''
                count += 1
                if vBCSTRet is not None:
                    vBCSTRet = vBCSTRet.text
                else:
                    vBCSTRet = ''
                if vICMSSTRet is not None:
                    vICMSSTRet = vICMSSTRet.text
                else:
                    vICMSSTRet = ''
                if vPMC is not None:
                    vPMC = vPMC.text
                else:
                    vPMC = ''
                listaca.append(f'{xml[0:44]}|{pos + 1}|{ean}|{reganvisa}|{cprod}|{descricao}|{ncm}|'
                               f'{qTrib}|{vUnTrib}|{vProd}|{vBCSTRet}|{vICMSSTRet}|{aliquota}|{vPMC}|{pmc12}|{pmc18}|'
                               f'{lcct}| '
                               f'{tipoproduto}|{fatorreducao}\n')
                print(f'{xml[0:44]}|{pos + 1}|{ean}|{reganvisa}|{cprod}|{descricao}|{ncm}|'
                      f'{qTrib}|{vUnTrib}|{vProd}|{vBCSTRet}|{vICMSSTRet}|{aliquota}|{vPMC}|{pmc12}|{pmc18}|{lcct}|'
                      f'{tipoproduto}|{fatorreducao}\n')
                with open(f'Nota_ComplementarCA_{periodo}.txt', 'w') as ncomplementar:
                    ncomplementar.writelines(listaca)
