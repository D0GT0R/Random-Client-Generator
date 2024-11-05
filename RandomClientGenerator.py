from datetime import date
import names
import random
import re
import xlrd
import xlsxwriter

###---Sanitize Input Logic---###
def sanitize_input(input_str):
    sanitized_str = re.sub(r'<script\b[^>]*>(.*?)</script>', '', input_str, flags=re.IGNORECASE)
    return sanitized_str

###---Read spreadsheet for address data---###
workbook1 = xlrd.open_workbook('AddressData.xls')
sheet = workbook1.sheet_by_name('Sheet1')

###---Check for/allow numeric values only for client quantity---###
while True:
    try:
        clientCountInt = int(sanitize_input(input('How many clients are needed?\n')))
    except ValueError:
        print('Error: Value must be numeric.')
        continue
    break

###---Enter Username (For file name)---###
userNameInput = sanitize_input(input('User requesting clients:\n'))

todaysDate = date.today().strftime("%b-%d-%y")

###---Create Excel File---###
workbook2 = xlsxwriter.Workbook(userNameInput + ' Clients ' + todaysDate + '.xlsx')
worksheet = workbook2.add_worksheet()

###---Function to set headers (needs work)---###
def set_headers():

    row = 0
    col = 0

    worksheet.set_column(0, 5, 25)

    worksheet.write(row, col, 'First Name')
    worksheet.write(row, col + 1, 'Last Name')
    worksheet.write(row, col + 2, 'Address')
    worksheet.write(row, col + 3, 'City')
    worksheet.write(row, col + 4, 'Zip Code')
    worksheet.write(row, col + 5, 'State')

###---Call Headers Function---###
set_headers()

row = 1
col = 0

###---For Loop to populate the Excel file with clients---###
for x in range(clientCountInt):

    randRow = random.randint(1, 10000)
    selectRow = int(randRow)

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

###---Close/Save the created Excel file---###
workbook2.close()

print('The list has been genrated with ' + str(clientCountInt) + ' clients.')
