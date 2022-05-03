from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter
from tabulate import tabulate
import os

wb = load_workbook('bank.xlsx')
ws = wb.active

def does_value_exist(value,col1,col2):
  count = 0
  uid = value
  for row in range(1,23):
    for col in range(col1,col2):
      char = get_column_letter(col)
      if (ws[char + str(row)].value) == uid:
        cell_row = row
        count +=1
        break
    if (ws[char + str(row)].value) == uid:
      break
  if count == 1:
    count = 0
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

def commandList(account_type):
  if account_type == "general user":
    print("-------------------------\n"
      "What do you want to do :\n"
      "(1) Deposit balance\n"
      "(2) Withdraw balance\n"
      "(3) Transfer balance\n"
      "(4) See Status\n"
      "(5) Exit\n"
      "-------------------------\n"
    )
  elif account_type == "admin":
    print(
      "-------------------------\n"
      "What do you want to do :\n"
      "(1) Add Data\n"
      "(2) Remove Data\n"
      "(3) Edit Data\n"
      "(4) Exit"
      "-------------------------\n")


def Deposit(cell_row,type,transferred):
  if type == "deposit":
    os.system("cls")
    deposit_amt = int(input("How much would you like to deposit?: "))
  elif type == "transfer":
    deposit_amt = transferred
  for row in range(cell_row,cell_row+1):
    for col in range(5,6):
      char = get_column_letter(col)
      previous_bal = ws[char + str(row)].value
      ws[char + str(row)].value = int(previous_bal)+deposit_amt
  new_bal = previous_bal + deposit_amt

  if type == "deposit":    
    print(f"Successfully deposited amount {deposit_amt}\n"
    f"Your new balance is {new_bal}\n")
  wb.save('bank.xlsx')
  
def Withdraw(cell_row,type):
  for row in range(cell_row,cell_row+1):
    for col in range(5,6):
      char = get_column_letter(col)
      previous_bal = ws[char + str(row)].value
      withdraw_amt = int(input(f"How much would you like to {type}?\n"
      f"maximum amount |{previous_bal}|: "))
      if withdraw_amt< previous_bal:
        ws[char + str(row)].value = int(previous_bal)-withdraw_amt
        new_bal = previous_bal - withdraw_amt
        pronoun = "withdrew" if type == "withdraw" else "transferred"
        print(
          f"Successfully {pronoun} amount {withdraw_amt}\n"
          f"Your new balance is {new_bal}\n")
      else:
        print(f"You dont have enough balance for the {type}")

  wb.save('bank.xlsx')
  return withdraw_amt

def Transfer(cell_row):
  receiver = int(input("Enter account number of the receiver: "))
  receiver_acc_no = does_value_exist(receiver,2,3)
  if receiver_acc_no != False:
    transfer = Withdraw(cell_row,"transfer")
    Deposit(receiver_acc_no,"transfer",transfer)
    

def available_function(account_type,cell_row):
  if account_type == "general user":
    commandList("general user")
    while True:
      command = input("command number: ")
      if command == "1":
        Deposit(cell_row,"deposit",0)
        commandList("general user")
      elif command == "2":
        os.system("cls")
        Withdraw(cell_row,"withdraw")
        commandList("general user")
      elif command == "3":
        os.system("cls")
        Transfer(cell_row)
        commandList("general user")
      elif command == "4":
        os.system("cls")
        print("User details\n")
        account_details(cell_row)
        commandList("general user")
      else:  
        break

  else:
    commandList("admin")
    
print(
  "---------------\n"
  "Welcome to bank\n"
  "---------------\n"
)

acount_type = " "

def main():
  uid = int(input("Enter you ID: "))
  cell_row = (does_value_exist(uid,1,2))
  if cell_row != False:
    pin =input("Enter pin: ")
    if (pin_validity(pin,cell_row)) == "valid":
      print("\nUser details")
      account_type = account_details(cell_row)
      available_function(account_type,cell_row)
      is_continue = input("Do you want to continue? Y/N: ").lower()
      if(is_continue == 'y'):
        os.system("cls")
        main()


  else:
    print("UID doesnt exists in the databse\n"
      "create new account or try another UID\n")

main()

