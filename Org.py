from xlrd import open_workbook

wb = open_workbook('tweets_spiderman_201809101957.xlsx')

from collections import defaultdict as ddict

treinamento=ddict(int)
teste=ddict(int)

dicts=[treinamento,teste]

# Ler tudo
for sheet,d in zip(wb.sheets(),dicts):
    values = []
    for row in range(sheet.nrows):
        col_value = []
        for col in range(sheet.ncols):
            value  = (sheet.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append(value)
        d[col_value[0]]+=1


# Organizar
for index in range(len(dicts)):
    l=[[i,dicts[index][i]] for i in dicts[index]]
    l=sorted(l, key = lambda x: int(x[1]))
    dicts[index]={item[0]:item[1] for item in l}



import xlsxwriter

# Criar o excel
file_name="spiderman_org2.xlsx"
workbook = xlsxwriter.Workbook(file_name)


# Formatação bold + caixa
cell_format = workbook.add_format()

cell_format.set_bold()
cell_format.set_center_across()
cell_format.set_border()



for sheet,d in zip(wb.sheets(),dicts):
    worksheet = workbook.add_worksheet(sheet.name)
    x=0
    for val in d:
        # Escrever no excel o item com o número de vezes que apareceu
        for i in range(d[val]):
            worksheet.write(x,0,val)
            x+=1

    worksheet.write(0,0,list(d.keys())[0],cell_format)
workbook.close()
