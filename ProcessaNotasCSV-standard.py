import csv
import xlsxwriter
import os
import shutil
#Colocar nome dos arquivos na lista, sem o .csv
nomeArquivos = []
caminho='C:\\Users\\SeverinoRibeiroNeto\\workspace\\Projetos Python\\ECIT\\leitura'
for (pathec, diretorio, arquivos) in os.walk(caminho):
    for arquivo in arquivos:
        nomeArquivos.append(str(caminho+'\\'+arquivo))



def transformaExcel(valores, nomeArquivo):
    workbook = xlsxwriter.Workbook(nomeArquivo +'.xlsx')
    bold = workbook.add_format({'bold':True})
    worksheet = workbook.add_worksheet()
    #Preenche Cabeçalho do arquivo    
    worksheet.write('A1', 'Matricula', bold)
    worksheet.write('B1', 'Nome', bold)
    worksheet.write('C1', 'Matemática', bold)
#    worksheet.write('D1', 'Sociologia', bold)
#    worksheet.write('E1', 'Fisica', bold)
    row = 1
    col = 0
    #[sala, matricula, primeiroNome, ultimoNome,notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia]
    #Fazer a mudança quando mudar a matéria.
    for (sala, matricula, primeiroNome, ultimoNome, notaMatematica) in valores:
        worksheet.write(row, col, matricula)
        worksheet.write(row, col+1, primeiroNome + ' ' + ultimoNome)
        worksheet.write(row, col+2, str(notaMatematica))
       # worksheet.write(row, col+3, str(notaSociologia))
        #worksheet.write(row, col+4, str(notaFisica))
        row += 1  
    workbook.close()

for nomeArquivo in nomeArquivos:
    #NotasDisciplinas
    notaLinguaPortuguesa = 0.0
    notaLinguaEspanhola = 0.0
    notaLinguaInglesa = 0.0
    notaEducacaoFisica = 0.0
    notaArtes = 0.0
    notaBiologia = 0.0
    notaQuimica = 0.0
    notaFisica = 0.0
    notaMatematica = 0.0
    notaHistoria = 0.0
    notaGeografia = 0.0
    notaSociologia = 0.0
    notaFilosofia = 0.0

    #Preencher com o nome do arquivo csv
    nota = 0
    alunos = []
    try:
        arquivo = open(nomeArquivo, 'r', encoding='utf8')
        leitor = csv.DictReader(arquivo)
    except Exception as Erro:
        print("Error! ", Erro )

    for linha in leitor:
        sala = linha['Class']
        matricula = linha['ZipGrade ID']
        primeiroNome = linha['First Name']
        ultimoNome = linha['Last Name']
        #Colocar as materias que vão ser utilizadas
        
        notaMatematica = (float(linha['Q1'])) + (float(linha['Q2'])) + (float(linha['Q3']))+ (float(linha['Q4']))+ (float(linha['Q5'])) + (float(linha['Q6']))+ (float(linha['Q7']))+ (float(linha['Q8']))+ (float(linha['Q9']))+ (float(linha['Q10']))
        #Salvando Alunos na lista
        #Não esquecer de mudar embaixo.
        alunos.append(tuple([sala, matricula, primeiroNome, ultimoNome,notaMatematica]))
    arquivo.close()
    nomeA = nomeArquivo.replace(caminho,'').replace('.csv','').replace('\\','')
    transformaExcel(alunos, nomeA)
    #Coloca o arquivo em excel com as notas e matérias
    shutil.move(nomeA+'.xlsx', caminho.replace('leitura','saida')+'\\'+nomeA+'.xlsx')
