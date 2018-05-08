import csv
from fpdf import FPDF
import os
import xlsxwriter
from tempfile import TemporaryFile

def transformaExcel(valores, nomeArquivo):
    workbook = xlsxwriter.Workbook(nomeArquivo +'.xlsx')
    worksheet = workbook.add_worksheet(valores[0][0])
    #Preenche Cabeçalho do arquivo    
    worksheet.write('A1', 'Matricula')
    worksheet.write('B1', 'Nome')
    worksheet.write('C1', 'Nota Lingua Portuguesa')
    worksheet.write('D1', 'Nota Lingua Espanhola')
    worksheet.write('E1', 'Nota Lingua Inglesa')
    worksheet.write('F1', 'Nota Educação Fisica')
    worksheet.write('G1', 'Nota Artes')
    worksheet.write('H1', 'Nota Biologia')
    worksheet.write('I1', 'Nota Quimica')
    worksheet.write('J1', 'Nota Fisica')
    worksheet.write('K1', 'Nota Matematica')
    worksheet.write('L1', 'Nota Historia')
    worksheet.write('M1', 'Nota Geografia')
    worksheet.write('N1', 'Nota Sociologia')
    worksheet.write('O1', 'Nota Filosofia')
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


# Limites das Disciplinas
linguaPortuguesa = 6
linguaEspanhola = 11
linguaInglesa = 16
educacaoFisica = 21
artes = 26
biologia = 31
quimica = 36
fisica = 41
matematica = 46
historia = 51
geografia = 56
sociologia = 61
filosofia = 66
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


nomeArquivo = 'testeStandard.csv'

arquivo = open(nomeArquivo, 'r')
leitor = csv.DictReader(arquivo)
nota = 0
alunos = []
for linha in leitor:
    sala = linha['Class']
    matricula = linha['ZipGrade ID']
    primeiroNome = linha['First Name']
    ultimoNome = linha['Last Name']

    # Lingua Portuguesa
    notaLinguaPortuguesa = float(linha['Q1']) + float(linha['Q2']) + float(linha['Q3']) + float(linha['Q4']) + float(linha['Q5'])
    # Lingua Espanhola
    notaLinguaEspanhola = float(linha['Q6']) + float(linha['Q7']) + float(linha['Q8']) + float(linha['Q9']) + float(linha['Q10'])
    # Lingua inglesa
    notaLinguaInglesa = float(linha['Q11']) + float(linha['Q12']) + float(linha['Q13']) + float(linha['Q14']) + float(linha['Q15'])
    # Educação Fisica
    notaEducacaoFisica = float(linha['Q16']) + float(linha['Q17']) + float(linha['Q18']) + float(linha['Q19']) + float(linha['Q20'])
    # Artes
    notaArtes = float(linha['Q21']) + float(linha['Q22']) + float(linha['Q23']) + float(linha['Q24']) + float(linha['Q25'])
    # Biologia
    notaBiologia = float(linha['Q26']) + float(linha['Q27']) + float(linha['Q28']) + float(linha['Q29']) + float(linha['Q30'])
    # Quimica
    notaQuimica = float(linha['Q31']) + float(linha['Q32']) + float(linha['Q33']) + float(linha['Q34']) + float(linha['Q35'])
    # Fisica
    notaFisica = float(linha['Q36']) + float(linha['Q37']) + float(linha['Q38']) + float(linha['Q39']) + float(linha['Q40'])
    # Matematica
    notaMatematica = float(linha['Q41']) + float(linha['Q42']) + float(linha['Q43']) + float(linha['Q44']) + float(linha['Q45'])
    # Historia
    notaHistoria = float(linha['Q46']) + float(linha['Q47']) + float(linha['Q48']) + float(linha['Q49']) + float(linha['Q50'])
    # Geografia
    notaGeografia = float(linha['Q51']) + float(linha['Q52']) + float(linha['Q53']) + float(linha['Q54']) + float(linha['Q55'])
    # Sociologia
    notaSociologia = float(linha['Q56']) + float(linha['Q57']) + float(linha['Q58']) + float(linha['Q59']) + float(linha['Q60'])
    # Filosofia
    notaFilosofia = float(linha['Q61']) + float(linha['Q62']) + float(linha['Q63']) + float(linha['Q64']) + float(linha['Q65'])
    #Salvando Alunos na lista
    alunos.append(tuple([sala, matricula, primeiroNome, ultimoNome,notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia]))

transformaExcel(alunos, 'testeAno')