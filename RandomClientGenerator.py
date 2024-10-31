import random_address
import xlsxwriter
import names
import json
from datetime import date



inp = input('How many clients are needed?\n')
clientCount = int(inp)

print(inp + ' clients are needed.')

today = date.today()

todayDate = today.strftime("%b-%d-%y")

workbook = xlsxwriter.Workbook('Clients ' + todayDate + '.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

worksheet.write(row, col, 'First Name')
worksheet.write(row, col + 1, 'Last Name')
worksheet.write(row, col + 2, 'Address 1')
worksheet.write(row, col + 3, 'Address 2')
worksheet.write(row, col + 4, 'City')
worksheet.write(row, col + 5, 'State')
worksheet.write(row, col + 6, 'Zipcode')

row = 1

for x in range(clientCount):
    iterationValue = (x + 1)
    d1 = random_address.real_random_address_by_state('GA')
    s1 = json.dumps(d1)
    d2 = json.loads(s1)

    clientFirst = names.get_first_name()
    clientLast = names.get_last_name()

    add1 = d2["address1"]
    add2 = d2["address2"]
    city = d2["city"]
    state = d2["state"]
    zipCode = d2["postalCode"]

    worksheet.write(row, col, clientFirst)
    worksheet.write(row, col + 1, clientLast)
    worksheet.write(row, col + 2, add1)
    worksheet.write(row, col + 3, add2)
    worksheet.write(row, col + 4, city)
    worksheet.write(row, col + 5, state)
    worksheet.write(row, col + 6, zipCode)

    row += 1

print('The list has been genrated')

workbook.close()