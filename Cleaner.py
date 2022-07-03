from openpyxl import load_workbook
from tkinter import *
from tkcalendar import Calendar
import datetime
import requests
def ColorCheck(column, line, sh):
    if ((sh[column+str(line)].fill.start_color.index != sh[column+str(line+1)].fill.start_color.index) and (sh[column+str(line+1)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line+1)].fill.start_color.index != 'FFC2C0C4')):
        return(sh[column+"2"].value + ': ' + sh[column+str(line+1)].value)

def EmptyCheck(column, line, sh):
    if sh[column +str(line)].fill.start_color.index == 'FFFFFFFF' or sh[column+str(line)].fill.start_color.index == 'FFC2C0C4':
        return(sh[column+"2"].value)

def LeavCheck(column, line, sh):
    if (sh[column+str(line)].fill.start_color.index != 'FFC2C0C4' and sh[column+str(line)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line+1)].fill.start_color.index == 'FFC2C0C4' or sh[column+str(line+1)].fill.start_color.index == 'FFFFFFFF'):
        return(sh[column+"2"].value)

def grad_date():
    
    listbox.delete(0,END)
    listbox2.delete(0,END)
    listbox3.delete(0,END)
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
    listbox2.insert(END, "Dnes Najizdi")
    listbox3.insert(END, "Dnes nikdo nenajede")
    for column in alphabet:
        listbox.insert(END, LeavCheck(column, line, sh))
        listbox2.insert(END, ColorCheck(column, line, sh))
        listbox3.insert(END, EmptyCheck(column, line, sh))
    

now = datetime.datetime.now()
"""
URL= "https://docs.google.com/spreadsheets/d/17gqV91AY6qDqFDWn2Y_YXSkr_h43csFG4hMUJSWKeQ0/export?format=xlsx"
response = requests.get(URL)
open("Obsazenost.xlsx", "wb").write(response.content)
"""
year = now.year
if year != 2022:
    print("Stary script! nutna obnova")
    exit(1)



wb = load_workbook("Obsazenost.xlsx", data_only=True)
sh = wb["Rozpis pokoju"]



window = Tk()
window.title("(u)KLID")
window.geometry("600x400")
window.configure(bg='lightgray')
cal = Calendar(window, selectmode = 'day',
               year = now.year, month = now.month,
               day = now.day, locale="cs_CZ")

cal.pack(pady = 5, fill= BOTH)
Button(window, text = "Get Date", command = grad_date).pack(pady = 20)
listbox = Listbox(window)
listbox2 = Listbox(window)
listbox3 = Listbox(window)
listbox.pack(expand=True, side="left", padx=5, fill=BOTH)
listbox2.pack(expand=True, side="left", padx=5, fill=BOTH)
listbox3.pack(expand=True, side="left", padx=5, fill=BOTH)
 






"""
print('--------------------------dnes odjizdi----------------------')
LeavCheck('e', line, sh)
LeavCheck('f', line, sh)
LeavCheck('g', line, sh)
LeavCheck('h', line, sh)
LeavCheck('i', line, sh)
LeavCheck('j', line, sh)
LeavCheck('k', line, sh)
LeavCheck('l', line, sh)
LeavCheck('m', line, sh)
LeavCheck('n', line, sh)
LeavCheck('o', line, sh)
LeavCheck('p', line, sh)
print('-------------------------dnes odjizdi-----------------------')
print()
print('-------------------------dnes najedou-----------------------')
ColorCheck('e', line, sh)
ColorCheck('f', line, sh)
ColorCheck('g', line, sh)
ColorCheck('h', line, sh)
ColorCheck('i', line, sh)
ColorCheck('j', line, sh)
ColorCheck('k', line, sh)
ColorCheck('l', line, sh)
ColorCheck('m', line, sh)
ColorCheck('n', line, sh)
ColorCheck('o', line, sh)
ColorCheck('p', line, sh)
print('-------------------------dnes najedou-----------------------')
print()
print('---------------------nikdo dnes nenajizdi--------------------')
EmptyCheck('e', line+1, sh)
EmptyCheck('f', line+1, sh)
EmptyCheck('g', line+1, sh)
EmptyCheck('h', line+1, sh)
EmptyCheck('i', line+1, sh)
EmptyCheck('j', line+1, sh)
EmptyCheck('k', line+1, sh)
EmptyCheck('l', line+1, sh)
EmptyCheck('m', line+1, sh)
EmptyCheck('n', line+1, sh)
EmptyCheck('o', line+1, sh)
EmptyCheck('p', line+1, sh)
print('----------------------nikdo dnes nenajizdi-------------------')
input('stisknete libovolnou klavesu pro ukonceni programu')
"""
window.mainloop()