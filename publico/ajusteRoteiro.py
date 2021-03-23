from openpyxl import load_workbook
import openpyxl
from openpyxl.worksheet import worksheet
from datetime import date
from openpyxl.styles import PatternFill

# Variaveis
carro = []
com = []
fim = []
notas1 = []
notas2 = []
notas3 = []
notas4 = []
notas5 = []
notas6 = []
feitas = []
hoje = date.today().strftime("%d.%m.%Y")
roteiro =str("C://Users/user/Local/NomeFinaldoArquivo"))

#funções
def check():
    for i in range(0, len(carro)):
        while True:
            try:
                x = int(input(f'Em qual nota o {carro[i]} está? '))
                feitas.append(x)
                break
            except:
                continue
    return feitas

#salva o arquivo modificado
def salvar():
    openpyxl.Workbook.save(planilha, filename=roteiro)

#abrir planilha e remover abas padrao excell vazias
planilha = load_workbook(filename="C://Users/user/Local/NomeFinaldoArquivo")
try:
    aba1 = planilha['Plan2']
    planilha.remove(aba1)
except:
    print()

try:
    aba2 = planilha['Plan3']
    planilha.remove(aba2)
except:
    print()

#declarar aba ativa
aba = planilha['Plan1']

#pegar dados dos carros para começar formatação
for i in range(0, 300):
    nNota = aba.cell(row=i+1, column=1).value
    placa = aba.cell(row=i+1, column=2).value
    if nNota == 1 :
        carro.append(placa)
        com.append(i+1)
    if nNota == None:break
       
for i in range(0, 300):
    nNota = aba.cell(row=i+1, column=1).value
    if nNota == 1 and i+1 != 2:
        fim.append(i+1)
    elif nNota == None:
        fim.append(i+1)
        break

#formatação da planilha / remoção dos dados inuteis
for i in range(0, 300):
    descr = aba.cell(row=i+1, column=13).value
    if descr == None:
        break
    else:
        aba.cell(row=i+1, column=13).value = None

worksheet.Worksheet.delete_cols(aba, 3)
worksheet.Worksheet.delete_cols(aba, 4)
worksheet.Worksheet.delete_cols(aba, 4)
worksheet.Worksheet.delete_cols(aba, 7)
worksheet.Worksheet.delete_cols(aba, 7)

#formatação da planilha / separando carros    
ws1 = planilha.create_sheet('SheetA')
try:
    ws1.title = carro[0]
    lin = 0
    col = 0
    for i in range(com[0], fim[0]):
        lin = lin+1
        col = 0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws1.cell(row=lin, column=col).value = planP.value           
except:
    planilha.remove(ws1)

ws2 = planilha.create_sheet('SheetA')
try:   
    ws2.title = carro[1]
    lin = 0
    col = 0
    for i in range(com[1], fim[1]):
        lin = lin+1
        col = 0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws2.cell(row=lin, column=col).value = planP.value
        
except:
    planilha.remove(ws2)

ws3 = planilha.create_sheet('SheetA')
try:    
    ws3.title = carro[2]
    lin = 0
    col = 0
    for i in range(com[2], fim[2]):
        lin = lin+1
        col = 0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws3.cell(row=lin, column=col).value = planP.value
except:
    planilha.remove(ws3)

ws4 = planilha.create_sheet('SheetA')
try:    
    ws4.title = carro[3]
    lin =0
    col = 0
    for i in range(com[3], fim[3]):
        lin = lin+1
        col = 0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws4.cell(row=lin, column=col).value = planP.value
except:
    planilha.remove(ws4)

ws5 = planilha.create_sheet('SheetA')
try:    
    ws5.title = carro[4]
    lin =0
    col = 0
    for i in range(com[4], fim[4]):
        lin = lin+1
        col =0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws5.cell(row=lin, column=col).value = planP.value
except:
    planilha.remove(ws5)

ws6 = planilha.create_sheet('SheetA')
try:    
    ws6.title = carro[5]
    lin =0
    col = 0
    for i in range(com[5], fim[5]):
        lin = lin+1
        col = 0
        for j in range(1 ,10):
            col = col+1
            planP = aba.cell(row=i, column=j)
            ws6.cell(row=lin, column=col).value = planP.value
except:
    planilha.remove(ws6)

#pegando novos com e fim para formatação da planilha final
com.clear()
fim.clear()

for i in range(0, 300):
    nNota = ws1.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)


for i in range(0, 300):
    nNota = ws1.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

for i in range(0, 300):
    nNota = ws2.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)
    if nNota == None:break

for i in range(0, 300):
    nNota = ws2.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

for i in range(0, 300):
    nNota = ws3.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)
    if nNota == None:break

for i in range(0, 300):
    nNota = ws3.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

for i in range(0, 300):
    nNota = ws4.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)
    if nNota == None:break

for i in range(0, 300):
    nNota = ws4.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

for i in range(0, 300):
    nNota = ws5.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)
    if nNota == None:break

for i in range(0, 300):
    nNota = ws5.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

for i in range(0, 300):
    nNota = ws6.cell(row=i+1, column=1).value
    if nNota == 1 :
        com.append(i+1)
    if nNota == None:break

for i in range(0, 300):
    nNota = ws6.cell(row=i+1, column=1).value
    if nNota == None:
        fim.append(i+1)
        break

#Formatação final da planilha principal
lin = 0
col = 0

for i in range(com[0], fim[0]):
    lin = lin+1
    col = 0
    for j in range(1, 10):
        col = col+1
        plan = ws1.cell(row=i, column=j)
        aba.cell(row=lin+1, column=col).value = plan.value
col = 0
for j in range(1,10):
    col = col+1
    aba.cell(row=lin+2, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
try:
    for i in range(com[1], fim[1]):
        lin = lin+1
        col = 0
        for j in range(1, 10):
            col = col+1
            plan = ws2.cell(row=i, column=j)
            aba.cell(row=lin+2, column=col).value = plan.value

    col = 0
    for j in range(1,10):
        col = col+1
        aba.cell(row=lin+3, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
except:
    pass

try:
    for i in range(com[2], fim[2]):
        lin = lin+1
        col = 0
        for j in range(1, 10):
            col = col+1
            plan = ws3.cell(row=i, column=j)
            aba.cell(row=lin+3, column=col).value = plan.value

    col = 0
    for j in range(1,10):
        col = col+1
        aba.cell(row=lin+4, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
except:
    pass
try:
    for i in range(com[3], fim[3]):
        lin = lin+1
        col = 0
        for j in range(1, 10):
            col = col+1
            plan = ws4.cell(row=i, column=j)
            aba.cell(row=lin+4, column=col).value = plan.value

    col = 0
    for j in range(1,10):
        col = col+1
        aba.cell(row=lin+5, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
except:
    pass

try:
    for i in range(com[4], fim[4]):
        lin = lin+1
        col = 0
        for j in range(1, 10):
            col = col+1
            plan = ws5.cell(row=i, column=j)
            aba.cell(row=lin+5, column=col).value = plan.value

    col = 0
    for j in range(1,10):
        col = col+1
        aba.cell(row=lin+6, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
except:
    pass

try:
    for i in range(com[5], fim[5]):
        lin = lin+1
        col = 0
        for j in range(1, 10):
            col = col+1
            plan = ws6.cell(row=i, column=j)
            aba.cell(row=lin+6, column=col).value = plan.value

    col = 0
    for j in range(1,10):
        col = col+1
        aba.cell(row=lin+7, column=col).fill = PatternFill(fgColor="808080", fill_type = "solid")
except:pass

#Salvando numero das notas do carro
try:   
    for i in range(com[0], fim[0]):
        notas1.append(ws1.cell(row=i, column=3).value)
        

except:pass

try:   
    for i in range(com[1], fim[1]):
        notas2.append(ws2.cell(row=i, column=3).value)
        

except:pass

try:   
    for i in range(com[2], fim[2]):
        notas3.append(ws3.cell(row=i, column=3).value)
        

except:pass

try:   
    for i in range(com[3], fim[3]):
        notas4.append(ws4.cell(row=i, column=3).value)
        

except:pass

try:   
    for i in range(com[4], fim[4]):
        notas5.append(ws5.cell(row=i, column=3).value)
        

except:pass

try:   
    for i in range(com[5], fim[5]):
        notas6.append(ws6.cell(row=i, column=3).value)
        

except:pass

salvar()
