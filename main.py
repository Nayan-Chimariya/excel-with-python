from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter
from tabulate import tabulate

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
    for col in range(6,7):
      char = get_column_letter(col)
      if str(ws[char + str(row)].value) == (pin):
        return "valid"
      else:
        return "invalid"  

def account_details(cell_row):
  title_list=[]
  details_list =[]

  for row in range(1,2):
    for col in range(2,6):
      char = get_column_letter(col)
      title_list.append((ws[char + str(row)].value))

  for row in range(cell_row,cell_row+1):
    for col in range(2,6):
      char = get_column_letter(col)
      details_list.append((ws[char + str(row)].value))

  combined_list = [list(l) for l in zip(title_list, details_list)]
  
  print(tabulate(combined_list))
  return details_list[2]

def commandList():
  print("-------------------------\n"
    "What do you want to do :\n"
    "(1) Deposit balance\n"
    "(2) Withdraw balance\n"
    "(3) Transfer balance\n"
    "(4) Exit\n"
    "-------------------------\n"
  )

def Deposit(cell_row):
  deposit_amt = int(input("How much would you like to deposit?: "))
  for row in range(cell_row,cell_row+1):
    for col in range(5,6):
      char = get_column_letter(col)
      previous_bal = ws[char + str(row)].value
      ws[char + str(row)].value = int(previous_bal)+deposit_amt
  new_bal = previous_bal + deposit_amt
      
  print(f"Successfully deposited amount {deposit_amt}\n"
  f"Your new balance is {new_bal}\n")
  wb.save('bank.xlsx')
  commandList()
  
def Withdraw(cell_row):
  for row in range(cell_row,cell_row+1):
    for col in range(5,6):
      char = get_column_letter(col)
      previous_bal = ws[char + str(row)].value
      withdraw_amt = int(input("How much would you like to withdraw?\n"
      f"maximum amount |{previous_bal}|: "))
      if withdraw_amt< previous_bal:
        ws[char + str(row)].value = int(previous_bal)-withdraw_amt
        new_bal = previous_bal - withdraw_amt
        print(
          f"Successfully withdrew amount {withdraw_amt}\n"
          f"Your new balance is {new_bal}\n")
      else:
        print("You dont have enough balance for the withdrawl")

  wb.save('bank.xlsx')
  commandList()

def available_function(account_type,cell_row):
  if account_type == "general user":
    commandList()
    while True:
      command = input("command number: ")
      if command == "1":
        Deposit(cell_row)
      elif command == "2":
        Withdraw(cell_row)
      elif command == 3:
        print("command 3")
      else:
        break

  else:
    print("What do you want to do :\n"
    "(1) Add Data\n"
    "(2) Remove Data\n"
    "(3) Edit Data\n")



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
  if (pin_validity(pin,cell_row)) == "valid":
    print("\nUser details")
    account_type = account_details(cell_row)
    available_function(account_type,cell_row)
    print("test")

else:
  print("UID doesnt exists in the databse\n"
    "create new account or try another UID\n")


#work in progress 
'''acc_no = int(input("Enter you account number: "))

#checking through UID column in excel sheet for account type

for row in range(1,21):
  for col in range(1,2):
    char = get_column_letter(col)'''
