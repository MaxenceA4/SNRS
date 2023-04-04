import os
import sys
import xlwings as xw
import shutil
import time

# Start Excel
xw.App(visible=True, add_book=False)

# Create a new workbook for each Serial Number in the first book in column C, and copy some data from the first book to the new workbook
# Open the first book
wb = xw.Book('Jauge pression.xlsx')
# Open the sheet "inventaire jauges de pression"
sht = wb.sheets['inventaire jauges de pression']
# Count the number of rows in the sheet jumping over the empty cells
last_row = sht.range('C2').end('down').row
print('Number of rows in the sheet: {}'.format(last_row))
for row in range(2, last_row + 1):
    print('Row: {}'.format(row))
    # store all the values of the row in a list\
    row_values = sht.range('A{}:K{}'.format(row, row)).value
    print('Row values: {}'.format(row_values))
    # if the 3rd value of the list is empty, then skip the row
    if row_values[2] == None:
        continue
    # if the 3rd value of the list is not empty, then create a new workbook for the Serial Number
    else:
        # Si la 3eme valeur de la liste finit par ".0", alors supprimer les 2 derniers caracteres
        if str(row_values[2]).endswith('.0'):
            row_values[2] = str(row_values[2])[:-2]
        else:
            row_values[2] = row_values[2]
        # Create a new workbook for the Serial Number
        shutil.copyfile('protected/template.xlsx', 'test/{}.xlsx'.format(row_values[2]))
        print('Created workbook for Serial Number {}'.format(row_values[2]))
        # Open the new workbook
        wb_new = xw.Book('test/{}.xlsx'.format(row_values[2]))
        # Open the sheet "inventaire jauges de pression"
        sht_new = wb_new.sheets['Valeur']

        ##########################################
        # Copy some data from the list to the new workbook
        # 1st item of list = B1
        sht_new.range('B1').value = row_values[0]
        # 2nd item of list = B2
        sht_new.range('B2').value = row_values[1]
        # 3rd item of list = B3
        sht_new.range('B3').value = row_values[2]
        # 4th item of list = skip
        # 5th item of list = B4
        sht_new.range('B4').value = row_values[4]
        # 6th item of list = B6
        sht_new.range('B6').value = row_values[5]
        # 7th item of list = B7
        sht_new.range('B7').value = row_values[6]
        # 8th item of list = B19
        sht_new.range('B18').value = row_values[7]
        # 9th item of list = skip
        # 10th item of list = B20
        sht_new.range('B19').value = row_values[9]
        # 11th item of list = B17
        sht_new.range('B17').value = row_values[10]
        ##########################################
        # Save the new workbook
        wb_new.save()
        # Close the new workbook
        wb_new.close()

        # next row
        continue
# Close the first workbook
wb.close()
#Exit Excel
xw.App.quit()

print('Done')
