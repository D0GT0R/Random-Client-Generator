import time
import xlsxwriter
import xlrd
import names
import random
from datetime import date

workbook = xlrd.open_workbook('AddressData.xls')
sheet = workbook.sheet_by_name('Sheet1')

inp = input('How many clients are needed?\n')
clientCount = int(inp)

today = date.today()

todayDate = today.strftime("%b-%d-%y")

workbook = xlsxwriter.Workbook('Clients ' + todayDate + '.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

worksheet.write(row, col, 'First Name')
worksheet.write(row, col + 1, 'Last Name')
worksheet.write(row, col + 2, 'Address 1')
worksheet.write(row, col + 3, 'City')
worksheet.write(row, col + 4, 'State')
worksheet.write(row, col + 5, 'Zipcode')

row = 1

for x in range(clientCount):

    randRow = random.randint(1, 820)
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

print('The list has been genrated with ' + inp + ' clients.')

workbook.close()