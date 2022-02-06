from tkinter import filedialog
from unicodedata import name
import urllib.request
import openpyxl
from bs4 import BeautifulSoup
import xlsxwriter
import time
from tkinter import *

"""
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
row2=0
compteurpasvide = 0

for cell in sheet["C"]:
  row = row+1
  print(row)
  if type(cell.value) is str:
    if cell.value != "Link":
      try:
        print(cell.value)
        soup = BeautifulSoup(urllib.request.urlopen(cell.value), 'lxml')
        tableau1 = soup('table', {"class" : "BordCollapseYear2"})[0] # premier tableau contenant capitalisation et PER
        tableau2 = soup('table', {"class" : "BordCollapseYear2"})[1]

        capitalisationtemp = Convert(tableau1.findAll('tr')[5].get_text(" "))
        capitalisation = capitalisationtemp[5:13]
        print(capitalisationtemp)
        for i in range(8):
          if capitalisation[i] != '\n':
            sheet[chr(ord("M")+i)][row-1].value = capitalisation[i]


        pertemp = Convert(tableau1.findAll('tr')[3].get_text(" "))
        per = pertemp[3:11]
        print(per)
        for i in range(8):
          try:
            if per[i] != '\n':
              sheet[chr(ord("E")+i)][row-1].value = per[i]
          except IndexError:
            print("Problème avec le nombre de données")


        bnatemp = Convert(tableau2.findAll('tr')[8].get_text(" "))
        bna = bnatemp[4:12]
        print(bna)
        for i in range(8):
          if bna[i] != '\n':
            sheet[col('U', i)][row-1].value = bna[i]
          
      except IndexError:
        print("Erreur dans le nombre de données récupérées, vérifiez que la page est valide")
        pass




for cell in sheet["D"]:
  try:
    row2 = row2+1
    print(row2)
    if type(cell.value) is str:
      if cell.value != "Link Beta":
        print(cell.value)
        soup = BeautifulSoup(urllib.request.urlopen(cell.value), 'lxml')
        issoutest = soup('span', {"class" : "mod-ui-data-list__value"})[4].text
        issouconvert = Convert(issoutest)
        print(issouconvert[0])
        sheet["AC"][row2-1].value = issouconvert[0]
  except IndexError:
    print("Erreur dans le nombre de données récupérées, vérifiez que la page est valide")
    pass



wb.save(namefile)
"""
#Partie GUI

def openFile():
  filepath= filedialog.askopenfilename()
  print(filepath)
  return filepath

root = Tk()
root.minsize(550, 400)
root.title("Financial Web Scraping Project")
frame = Frame(root, bg="#1e1e1e")


Label(frame, text="FINANCIAL WEB SCRAPING", bg="#1e1e1e", font=("Arial Bold", 18,), fg= "white" ).grid(row=0, column=0, ipadx=20, ipady=20)
button = Button(frame, text="Ouvrir votre fichier Excel", command=openFile)
button.grid(row=1, column=0)


frame.pack()
frame.place(relx = 0,
                   rely = 0,
                   anchor = 'nw')
root.configure(background='#1e1e1e')
root.mainloop()