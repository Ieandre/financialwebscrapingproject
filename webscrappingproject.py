from tkinter import filedialog
from unicodedata import name
import urllib.request
import openpyxl
from bs4 import BeautifulSoup
import xlsxwriter
import time
from tkinter import *

# Partie logique

def scraping(filepath):
  timestr = time.strftime("%d-%m-%Y_%H-%M-%S")

  namefile = str(timestr)+'.xlsx'



  wb = openpyxl.load_workbook(filepath)
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
          textbox.insert(END,"Erreur dans le nombre de données récupérées à la ligne "+str(row)+" vérifiez que la page est valide \n")
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
      textbox.insert(END,"Erreur dans le nombre de données récupérées pour le Beta, vérifiez que la page à la ligne " +str(row)+ " est valide \n")
      pass

  textbox.insert(END,"Fin du script, fichier sauvegardé")
  wb.save(namefile)


#Partie GUI

def openFile():
  global filepath
  filepath= filedialog.askopenfilename()
  #textbox.insert(1.0, "Fichier récuperé : " + str(filepath) + "\n")
  textbox.insert(END,"Fichier récuperé, cliquez sur le bouton pour lancer le scraping (peut prendre plusieurs minutes à s'éxecuter) \n" )



root = Tk()
root.minsize(600, 700)
root.title("Financial Web Scraping Project Beta")
root.iconbitmap('scraper.ico')
frame = Frame(root, bg="#121212")


Label(frame, text="WEB SCRAPING PROJECT", bg="#121212", font=("Arial Bold", 18,), fg= "#e3e3e3" ).grid(row=0, column=0, ipadx=20, ipady=20)
button = Button(frame, text="Ouvrir votre fichier Excel", command=openFile, bg='#bb86fc', fg="#e3e3e3", bd=0)
button.grid(row=2, column=0, pady=(20,0))

button2 = Button(frame, text="Executer le script", command=lambda:scraping(filepath), bg='#bb86fc', fg="#e3e3e3",bd=0)
button2.grid(row=3, column=0,pady=(20,0))

textbox= Text(frame, height=30, width=60, bg="#1e1e1e", fg='#e3e3e3')
textbox.grid(row=1, column=0)





frame.pack()
frame.place(relx = 0.5,
                   rely = 0.5,
                   anchor = 'center')
root.configure(background='#121212')
root.mainloop()