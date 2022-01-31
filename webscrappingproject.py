from unicodedata import name
import urllib.request
import openpyxl
from bs4 import BeautifulSoup
import xlsxwriter
import time

timestr = time.strftime("%d-%m-%Y_%H-%M-%S")

namefile = str(timestr)+'.xlsx'



wb = openpyxl.load_workbook(r'test.xlsx')
sheet = wb.active

def col(s, n):
  if len(s) == 1:
    c = ord(s) + n 
    if c > ord('Z'):
      return "A"+chr(c-26)
    return chr(c)
  else:
    return "A"+col(s[1:], n)

def Convert(string):
    li = list(string.split(" "))
    return li




row = 0
for cell in sheet["C"]:
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
          sheet[chr(ord("M")+i)][row-1].value = capitalisation[i]


      pertemp = Convert(tableau1.findAll('tr')[3].get_text(" "))
      per = pertemp[3:10]
      print(per)
      for i in range(7):
        if per[i] != '\n':
          sheet[chr(ord("E")+i)][row-1].value = per[i]

      bnatemp = Convert(tableau2.findAll('tr')[8].get_text(" "))
      bna = bnatemp[4:11]
      print(bna)
      for i in range(7):
        if bna[i] != '\n':
          print(i)
          sheet[col('U', i)][row-1].value = bna[i]



wb.save(namefile)
