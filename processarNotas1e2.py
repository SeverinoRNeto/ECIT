import csv
from fpdf import FPDF
import os
import xlsxwriter
from tempfile import TemporaryFile
#arquivos=['quiz-Simulado1AAgro-standard','quiz-Simulado1BAgro-standard','quiz-Simulado1CAgro-standard','quiz-Simulado1BMsi2-standard','quiz-Simulado1AMsi-standard','quiz-Simulado1CMsi-standard']
arquivos=['quiz-Simulado3AAgro-standard', 'quiz-Simulado3AMsi-standard', 'quiz-Simulado3BMsi-standard']
#arquivos = ['quiz-Simulado2AAGRO-standard', 'quiz-Simulado2AMsi-standard', 'quiz-Simulado2BAGRO-standard','quiz-Simulado2BMsi-standard', 'quiz-Simulado2CMSI-standard']
#excelNome= 'Simulado3AAgro'
#nomeArquivo = 'quiz-Simulado3AAgro-standard.csv'


def transformaExcel(valores, nomeArquivo):
    workbook = xlsxwriter.Workbook(nomeArquivo +'.xlsx')
    bold = workbook.add_format({'bold':True})
    worksheet = workbook.add_worksheet(valores[0][0])
    #Preenche Cabeçalho do arquivo    
    worksheet.write('A1', 'Matricula', bold)
    worksheet.write('B1', 'Nome', bold)
    worksheet.write('C1', 'Ling.Port', bold)
    worksheet.write('D1', 'Ling.Esp', bold)
    worksheet.write('E1', 'Ling.Ing', bold)
    worksheet.write('F1', 'Ed.Fisica',bold)
    worksheet.write('G1', 'Arte', bold)
    worksheet.write('H1', 'Biologia', bold)
    worksheet.write('I1', 'Quimica', bold)
    worksheet.write('J1', 'Fisica', bold)
    worksheet.write('K1', 'Matematica', bold)
    worksheet.write('L1', 'Historia', bold)
    worksheet.write('M1', 'Geografia', bold)
    worksheet.write('N1', 'Sociologia', bold)
    worksheet.write('O1', 'Filosofia', bold)
    worksheet.write('P1', 'Sala', bold)
    row = 2
    col = 0
    #[sala, matricula, primeiroNome, ultimoNome,notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia]
    for (sala, matricula, primeiroNome, ultimoNome, notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia) in valores:
        worksheet.write(row, col, matricula)
        worksheet.write(row, col+1, primeiroNome + ' ' + ultimoNome)
        worksheet.write(row, col+2, str(notaLinguaPortuguesa))
        worksheet.write(row, col +3, str(notaLinguaEspanhola))
        worksheet.write(row, col +4, str(notaLinguaInglesa))
        worksheet.write(row, col +5, str(notaEducacaoFisica))
        worksheet.write(row, col +6, str(notaArtes))
        worksheet.write(row, col +7, str(notaBiologia))
        worksheet.write(row, col +8, str(notaQuimica))
        worksheet.write(row, col +9, str(notaFisica))
        worksheet.write(row, col +10, str(notaMatematica))
        worksheet.write(row, col +11, str(notaHistoria))
        worksheet.write(row, col +12, str(notaGeografia))
        worksheet.write(row, col +13, str(notaSociologia))
        worksheet.write(row, col +14, str(notaFilosofia))
        worksheet.write(row, col +15, sala)
        row += 1  
    workbook.close()

for nomeArquivo in arquivos:
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
        arquivo = open(nomeArquivo + '.csv', 'r')
        leitor = csv.DictReader(arquivo)
    except:
        print("Error! " )

    for linha in leitor:
        sala = linha['Class']
        matricula = linha['ZipGrade ID']
        primeiroNome = linha['First Name']
        ultimoNome = linha['Last Name']

        # Lingua Portuguesa
        notaLinguaPortuguesa = (float(linha['Q1'])*2) + (float(linha['Q2'])*2) + (float(linha['Q2'])*2) + (float(linha['Q4'])*2) + (float(linha['Q5'])*2)
        # Lingua Espanhola
        notaLinguaEspanhola = (float(linha['Q6'])*2) + (float(linha['Q7'])*2) + (1*2) + (float(linha['Q9'])*2) + (float(linha['Q10'])*2)
        # Lingua inglesa
        notaLinguaInglesa = (float(linha['Q11'])*2) + (float(linha['Q12'])*2) + (float(linha['Q13'])*2) + (float(linha['Q14'])*2) + (float(linha['Q15'])*2)
        # Educação Fisica
        notaEducacaoFisica = (float(linha['Q16'])*2) + (float(linha['Q17'])*2) + (float(linha['Q18'])*2) + (float(linha['Q19'])*2) + (float(linha['Q20'])*2)
        # Artes
        notaArtes = (float(linha['Q21'])*2) + (float(linha['Q22'])*2) + (float(linha['Q23'])*2) + (float(linha['Q24'])*2) + (float(linha['Q25'])*2)
        # Biologia
        notaBiologia = (float(linha['Q26'])*2) + (float(linha['Q27'])*2) + (float(linha['Q28'])*2) + (float(linha['Q29'])*2) + (float(linha['Q30'])*2)
        # Quimica
        notaQuimica = (float(linha['Q31'])*2) + (float(linha['Q32'])*2) + (float(linha['Q33'])*2) + (float(linha['Q34'])*2) + (float(linha['Q35'])*2)
        # Fisica
        notaFisica = (float(linha['Q36'])*2) + (float(linha['Q37'])*2) + (float(linha['Q38'])*2) + (float(linha['Q39'])*2) + (float(linha['Q40'])*2)
        # Matematica
        notaMatematica = (float(linha['Q41'])*2) + (float(linha['Q42'])*2) + (float(linha['Q43'])*2) + (float(linha['Q44'])*2) + (float(linha['Q45'])*2)
        # Historia
        notaHistoria = (float(linha['Q46'])*2) + (float(linha['Q47'])*2) + (float(linha['Q48'])*2) + (float(linha['Q49'])*2) + (float(linha['Q50'])*2)
        # Geografia
        notaGeografia = (float(linha['Q51'])*2) + (float(linha['Q52'])*2) + (float(linha['Q53'])*2) + (float(linha['Q54'])*2) + (float(linha['Q55'])*2)
        # Sociologia/Filosofia no terceiro
        notaFilosofia = (float(linha['Q56'])*2) + (float(linha['Q57'])*2) + (float(linha['Q58'])*2) + (float(linha['Q59'])*2) + (float(linha['Q60'])*2)
        # Filosofia/Sociologia no terceiro
        notaSociologia = (float(linha['Q61'])*2) + (float(linha['Q62'])*2) + (float(linha['Q63'])*2) + (float(linha['Q64'])*2) + (float(linha['Q65'])*2)
        #Salvando Alunos na lista
        alunos.append(tuple([sala, matricula, primeiroNome, ultimoNome,notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia]))
    arquivo.close()
    transformaExcel(alunos, nomeArquivo)