import os
import sys
import xlwings as xw
import shutil
import time
# Create a new workbook for each Serial Number in the first book in column C, and copy some data from the first book to the new workbook
# Open the first book
wb = xw.Book('Jauge pression.xlsx')
# Open the sheet "inventaire jauges de pression"
sht = wb.sheets['inventaire jauges de pression']
# Count the number of rows in the sheet jumping over the empty cells
last_row = sht.range('C1').end('down').row
print('Number of rows in the sheet: {}'.format(last_row))
for row in range(2, last_row + 1):
    print('Row: {}'.format(row))
    # Store all the value of the row in a list
    row_values = sht.range('A{}:G{}'.format(row, row)).value
    print('Row values: {}'.format(row_values))
    # Print only the third value of the row
    print('Serial Number: {}'.format(row_values[2]))
    for value in row_values:
        print(value)
        

# # Create a new workbook for each unique Serial Number
# for unique_serial_number_and_row_number in unique_serial_numbers_and_row_numbers:
#     shutil.copyfile('protected/template.xlsx', 'test/{}.xlsx'.format(unique_serial_number_and_row_number[0]))
#     print('Created workbook for Serial Number {}'.format(unique_serial_number_and_row_number[0]))

