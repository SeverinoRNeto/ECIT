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
    for i in range(1, linguaPortuguesa):
        questao = 'Q' + str(i)
        notaLinguaPortuguesa = notaLinguaPortuguesa + int(linha[questao])
# Lingua Espanhola
    for i in range(linguaPortuguesa, linguaEspanhola):
        questao = 'Q' + str(i)
        notaLinguaEspanhola = notaLinguaEspanhola + int(linha[questao])
# Lingua inglesa
    for i in range(linguaEspanhola, linguaInglesa):
        questao = 'Q' + str(i)
        notaLinguaInglesa = notaLinguaInglesa + int(linha[questao])
# Educação Fisica
    for i in range(linguaInglesa, educacaoFisica):
        questao = 'Q' + str(i)
        notaEducacaoFisica = notaEducacaoFisica + int(linha[questao])
# Artes
    for i in range(educacaoFisica, artes):
        questao = 'Q' + str(i)
        notaArtes = notaArtes + int(linha[questao])
# Biologia
    for i in range(artes, biologia):
        questao = 'Q' + str(i)
        notaBiologia = notaBiologia + int(linha[questao])
# Quimica
    for i in range(biologia, quimica):
        questao = 'Q' + str(i)
        notaQuimica = notaQuimica + int(linha[questao])
# Fisica
    for i in range(quimica, fisica):
        questao = 'Q' + str(i)
        notaFisica = notaFisica + int(linha[questao])
# Matematica
    for i in range(fisica, matematica):
        questao = 'Q' + str(i)
        notaMatematica = notaMatematica + int(linha[questao])
# Historia
    for i in range(matematica, historia):
        questao = 'Q' + str(i)
        notaHistoria = notaHistoria + int(linha[questao])
# Geografia
    for i in range(historia, geografia):
        questao = 'Q' + str(i)
        notaGeografia = notaGeografia + int(linha[questao])
# Sociologia
    for i in range(geografia,sociologia):
        questao = 'Q' + str(i)
        notaSociologia = notaSociologia + int(linha[questao])
# Filosofia
    for i in range(sociologia, filosofia):
        questao = 'Q' + str(i)
        notaFilosofia = notaFilosofia + int(linha[questao])

    alunos.append(tuple([sala, matricula, primeiroNome, ultimoNome,notaLinguaPortuguesa, notaLinguaEspanhola, notaLinguaInglesa, notaEducacaoFisica, notaArtes, notaBiologia, notaQuimica, notaFisica, notaMatematica, notaHistoria, notaGeografia, notaSociologia, notaFilosofia]))

transformaExcel(alunos, 'primeiroAno')