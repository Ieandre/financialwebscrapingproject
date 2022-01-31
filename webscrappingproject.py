import urllib.request
import openpyxl
from bs4 import BeautifulSoup

wb = openpyxl.load_workbook(r'test.xlsx')
sheet = wb.active

def Convert(string):
    li = list(string.split(" "))
    return li

for cell in sheet["F"]:
  if type(cell.value) is str:
    if cell.value != "Link":
      print(cell.value)
      soup = BeautifulSoup(urllib.request.urlopen(cell.value), 'lxml')
      tableau1 = soup('table', {"class" : "BordCollapseYear2"})[0]
      tableau2 = soup('table', {"class" : "BordCollapseYear2"})[1]

      capitalisationtemp = Convert(tableau1.findAll('tr')[5].get_text(" "))
      capitalisation = capitalisationtemp[5:12]
      print(capitalisation)

      pertemp = Convert(tableau1.findAll('tr')[3].get_text(" "))
      per = pertemp[3:10]
      print(per)

      bnatemp = Convert(tableau2.findAll('tr')[8].get_text(" "))
      bna = bnatemp[4:11]
      print(bna)
