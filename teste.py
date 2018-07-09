import os
nomeArquivos=[]
for (pathec, diretorio, arquivos) in os.walk('C:\\Users\\SeverinoRibeiroNeto\\workspace\\Projetos Python\\ECIT\\leitura'):
    for arquivo in arquivos:
        nomeArquivos.append(str(arquivo).replace('.csv',''))
print(nomeArquivos)