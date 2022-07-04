from openpyxl import load_workbook
from tkinter import *
from tkcalendar import Calendar
import tkcap
from PIL import Image
import datetime
import requests
def ColorCheck(column, line, sh):
    #dnes najizdi
    if ((sh[column+str(line)].fill.start_color.index != sh[column+str(line+1)].fill.start_color.index) and (sh[column+str(line+1)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line+1)].fill.start_color.index != 'FFC2C0C4') and (sh[column+str(line+1)].fill.start_color.index != 'C2C0C4')):
        if (sh[column+str(line)].fill.start_color.index != sh[column+str(line+1)].fill.start_color.index) and (sh[column+str(line)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line)].fill.start_color.index != 'FFC2C0C4') and (sh[column+str(line)].fill.start_color.index != 'C2C0C4'):
            listbox.insert(END, sh[column+"2"].value)
        if (column == "i"):
            listbox2.insert(END, sh[column+"2"].value + ': ' + sh[column+str(line+1)].value + ' - Postele od sebe???')
        else:
            listbox2.insert(END, sh[column+"2"].value + ': ' + sh[column+str(line+1)].value)
    #dnes odjizdi
    elif (sh[column+str(line)].fill.start_color.index != 'FFC2C0C4' and sh[column+str(line)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line+1)].fill.start_color.index == 'FFC2C0C4' or sh[column+str(line+1)].fill.start_color.index == 'FFFFFFFF'):
        listbox.insert(END, sh[column+"2"].value)
def Screenshot():
    cap = tkcap.CAP(window)
    cap.capture('screenshot.jpg')
    im = Image.open(r"C:\Users\stofa\OneDrive\Documents\GitHub\Potok-clean-script\screenshot.jpg")
    left = 0
    top = 400
    right = 1000
    bottom = 900
    imEdited = im.crop((left, top, right, bottom))
    imEdited.show()
def grad_date():
    
    listbox.delete(0,END)
    listbox2.delete(0,END)
    alphabet = ["e","f","g","h","i","j","k","l","m","n","o","p"]
    DateSelected = cal.get_date()
    splitted = DateSelected.split(".")
    day = splitted[0]
    month = splitted[1]
    month = month[1:]
    monthJump = 0 
    dayJump = (int(day) * 2) - 2

    if month == '7':
        monthJump = 795

    elif month == '8':
        monthJump = 860

    elif month == '9':
        monthJump = 922

    elif month == '10':
        monthJump = 982

    elif month == '11':
        monthJump = 1044

    elif month == '12':
        monthJump = 1104

    else:
        print('neplatny mesic, spustte program znovu')
    line = dayJump + monthJump

    listbox.insert(END, "Dnes odjíždí")
    listbox2.insert(END, "Dnes Najíždí")
    for column in alphabet:
        ColorCheck(column, line, sh)
    

now = datetime.datetime.now()

URL= "https://docs.google.com/spreadsheets/d/17gqV91AY6qDqFDWn2Y_YXSkr_h43csFG4hMUJSWKeQ0/export?format=xlsx"
response = requests.get(URL)
open("Obsazenost.xlsx", "wb").write(response.content)

year = now.year
if year != 2022:
    print("Stary script! nutna obnova")
    exit(1)



wb = load_workbook("Obsazenost.xlsx", data_only=True)
sh = wb["Rozpis pokoju"]



window = Tk()
window.title("(u)KLID")
window.geometry("1100x900")
window.configure(bg='lightgray')
cal = Calendar(window, selectmode = 'day',
               year = now.year, month = now.month,
               day = now.day, locale="cs_CZ")

cal.pack(pady = 5, fill= BOTH)
Button(window, text = "Vypsat úklidy", command = grad_date).pack(pady = 20)

listbox = Listbox(window, font=("Arial", 15))
listbox2 = Listbox(window, font=("Arial", 15))
listbox.pack(expand=True, side="left", padx=5, fill=BOTH)
listbox2.pack(expand=True, side="left", padx=5, fill=BOTH)
 




window.mainloop()