import xmltodict
import os
import json
import pandas as pd

def getInfos(fileName, values):
    with open(f'nfs/{fileName}', 'rb') as XMLfile:
        dicFile = xmltodict.parse(XMLfile)
        # print(json.dumps(dicFile, indent=4))
        if "NFe" in dicFile:
            infoNF = dicFile["NFe"]["infNFe"]
        else:
            infoNF = dicFile["nfeProc"]["NFe"]["infNFe"]
        NFnumber = infoNF["ide"]["cNF"]
        emitter = infoNF["emit"]["xNome"]
        clientName = infoNF["dest"]["xNome"]
        addressClient = infoNF["emit"]["enderEmit"]
        if "vol" in infoNF["transp"]:
            weigth = infoNF["transp"]["vol"]["pesoB"]
        else:
            weigth = "Peso não informado"

        values.append([NFnumber, emitter, clientName, addressClient, weigth])

listFile = os.listdir("nfs")

column = ["NFnumber", "emitter", "clientName", "addressClient", "weigth"]
values = []

for file in listFile:
    getInfos(file, values)

table = pd.DataFrame(columns=column, data=values)
try:
    table.to_excel("NotasFiscais.xlsx", index=False)
    print("Arquivo XLSX Salvo com sucesso!" )
except:
    print("ERRO AO TENTAR CRIAR ARQUIVO")
