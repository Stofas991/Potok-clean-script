from openpyxl import load_workbook
import datetime
def ColorCheck(column, line, sh):
    if (sh[column+str(line)].fill.start_color.index != sh[column+str(line+1)].fill.start_color.index and sh[column+str(line+1)].fill.start_color.index != 'FFFFFFFF' and sh[column+str(line+1)].fill.start_color.index != 'FFC2C0C4'):
        print(sh[column+"2"].value + ': ' + sh[column+str(line+1)].value)

def EmptyCheck(column, line, sh):
    if sh[column +str(line)].fill.start_color.index == 'FFFFFFFF' or sh[column+str(line)].fill.start_color.index == 'FFC2C0C4':
        print(sh[column+"2"].value)

def LeavCheck(column, line, sh):
    if (sh[column+str(line)].fill.start_color.index != 'FFC2C0C4' and sh[column+str(line)].fill.start_color.index != 'FFFFFFFF') and (sh[column+str(line+1)].fill.start_color.index == 'FFC2C0C4' or sh[column+str(line+1)].fill.start_color.index == 'FFFFFFFF'):
        print(sh[column+"2"].value)

now = datetime.datetime.now()
year = now.year
if year != 2021:
    print("Stary script! nutna obnova")
    exit(1)



wb = load_workbook("Obsazenost.xlsx", data_only=True)
sh = wb["Rozpis pokoju"]
month = input("zadejte mesic:")
day = input("zadejte den:")


monthJump = 0 
dayJump = 0
dayJump = (int(day) * 2) - 2

if month == '1':
    monthJump = 800

elif month == '2':
    monthJump = 862

elif month == '3':
    monthJump = 918

elif month == '4':
    monthJump = 980

elif month == '5':
    monthJump = 1040

elif month == '6':
    monthJump = 1102

elif month == '7':
    monthJump = 1162

elif month == '8':
    monthJump = 1224

elif month == '9':
    monthJump = 1286

elif month == '10':
    monthJump = 1346

elif month == '11':
    monthJump = 1408

elif month == '12':
    monthJump = 1468

else:
    print('neplatny mesic, spustte program znovu')


line = dayJump + monthJump
print('---------------------dnes odjizdi---------------------------')
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
print('---------------------dnes odjizdi---------------------------')
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