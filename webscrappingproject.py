import urllib.request
import openpyxl
from bs4 import BeautifulSoup
import xlsxwriter

wb = openpyxl.load_workbook(r'test.xlsx')
sheet = wb.active

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.get_worksheet_by_name('Sheet1')
#worksheet = workbook.add_worksheet()
print(worksheet)

def Convert(string):
    li = list(string.split(" "))
    return li




row = 0
for cell in sheet["F"]:
  row = row+1
  print(row)
  if type(cell.value) is str:
    if cell.value != "Link":
      print(cell.value)
      soup = BeautifulSoup(urllib.request.urlopen(cell.value), 'lxml')
      tableau1 = soup('table', {"class" : "BordCollapseYear2"})[0] # premier tableau contenant capitalisation et PER
      tableau2 = soup('table', {"class" : "BordCollapseYear2"})[1]

      capitalisationtemp = Convert(tableau1.findAll('tr')[5].get_text(" "))
      capitalisation = capitalisationtemp[5:12]
      print(capitalisation)
      for i in range(7):
        if capitalisation[i] != '\n':
          # worksheet.write(row, ord('H')-ord('A')+1+i, capitalisation[i])
          sheet[chr(ord("N")+i)][row-1].value = capitalisation[i]


      pertemp = Convert(tableau1.findAll('tr')[3].get_text(" "))
      per = pertemp[3:10]
      print(per)
      for i in range(7):
        if per[i] != '\n':
          sheet[chr(ord("H")+i)][row-1].value = per[i]

      bnatemp = Convert(tableau2.findAll('tr')[8].get_text(" "))
      bna = bnatemp[4:11]
      print(bna)
      for i in range(7):
        if bna[i] != '\n':
          sheet[chr(ord("S")+i)][row-1].value = bna[i]


wb.save('test33.xlsx')