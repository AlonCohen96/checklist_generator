import openpyxl
from openpyxl.styles import PatternFill

#admin_file = 'admintestfile.xlsx'
admin_file = '#insert file path here'

wb_obj = openpyxl.load_workbook(admin_file)
sheet1 = wb_obj['COVID19_ZH_V5_Total']
sheet2 = wb_obj['counter']

closing_list = []

while True:
    answer = input('Enter ID of participants who succesfully participated, confirm with Enter, to stop enter <stop>: ')
    if answer == 'stop':
        break
    closing_list.append(answer)

counter = 11

for row in sheet1:

    ID = str(sheet1['D' + str(counter)].value)
    studienMappeGeneriert = sheet1['B' + str(counter)].value

    if ID in closing_list:
        if studienMappeGeneriert == 1:
            sheet1['L' + str(counter)] = 1
            print('Closed: ' + ID)
    counter += 1

wb_obj.save('#insert file path here')
#wb_obj.save('admintestfile.xlsx')

print(closing_list)