import re
import xlsxwriter
import xlrd
import names
import random
from datetime import date

def sanitize_input(input_str):
    sanitized_str = re.sub(r'<script\b[^>]*>(.*?)</script>', '', input_str, flags=re.IGNORECASE)
    return sanitized_str

workbook = xlrd.open_workbook('AddressData.xls')
sheet = workbook.sheet_by_name('Sheet1')

clientCountInput = input('How many clients are needed?\n')

clientCount = sanitize_input(clientCountInput)

clientCountInt = int(clientCount)

userNameInput = input('User requesting clients:\n')

userName = sanitize_input(userNameInput)

today = date.today()

todaysDate = today.strftime("%b-%d-%y")

workbook1 = xlsxwriter.Workbook(userName + ' Clients ' + todaysDate + '.xlsx')
worksheet = workbook1.add_worksheet()

row = 0
col = 0

worksheet.set_column(0, 5, 25)

worksheet.write(row, col, 'First Name')
worksheet.write(row, col + 1, 'Last Name')
worksheet.write(row, col + 2, 'Address')
worksheet.write(row, col + 3, 'City')
worksheet.write(row, col + 4, 'Zip Code')
worksheet.write(row, col + 5, 'State')

row = 1

for x in range(clientCountInt):

    randRow = random.randint(1, 2480)
    selectRow = int(randRow)

    iterationValue = (x + 1)
    
    add1 = sheet.cell(selectRow, 0).value
    city = sheet.cell(selectRow, 1).value
    zipCode = sheet.cell(selectRow, 2).value
    state = sheet.cell(selectRow, 3).value
    

    clientFirst = names.get_first_name()
    clientLast = names.get_last_name()

    worksheet.write(row, col, clientFirst)
    worksheet.write(row, col + 1, clientLast)
    worksheet.write(row, col + 2, add1)
    worksheet.write(row, col + 3, city)
    worksheet.write(row, col + 4, zipCode)
    worksheet.write(row, col + 5, state)

    row += 1

print('The list has been genrated with ' + clientCount + ' clients.')

workbook1.close()
