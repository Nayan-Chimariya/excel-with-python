from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('bank.xlsx')
ws = wb.active

def does_uid_exist(UID):
  count = 0
  uid = UID
  for row in range(1,21):
    for col in range(1,2):
      char = get_column_letter(col)
      if(ws[char + str(row)].value) == uid:
        cell_row = row
        count +=1
        break
  if count == 1:
    return cell_row
  else:
    return False

def pin_validity(pin,cell_row):
  for row in range(cell_row,cell_row+1):
    for col in range(5,6):
      char = get_column_letter(col)
      print(ws[char + str(row)].value)
      if str(ws[char + str(row)].value) == (pin):
        return "valid"
      else:
        return "invalid"  

print(
  "---------------\n"
  "Welcome to bank\n"
  "---------------\n"
)

acount_type = " "

uid = int(input("Enter you ID: "))
cell_row = (does_uid_exist(uid))
if cell_row != False:
  pin =input("Enter pin: ")
  print (pin_validity(pin,cell_row))
else:
  print("UID doesnt exists in the databse\n"
    "create new account or try another UID\n")


#work in progress 
'''acc_no = int(input("Enter you account number: "))

#checking through UID column in excel sheet for account type

for row in range(1,21):
  for col in range(1,2):
    char = get_column_letter(col)'''
