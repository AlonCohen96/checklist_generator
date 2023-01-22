import openpyxl

test_file = '#insert file path here'

wb_obj = openpyxl.load_workbook(test_file)
source_sheet = wb_obj.active

wb_massenversand_template = openpyxl.load_workbook('dependencies/massenversand_template.xlsx')
output_sheet = wb_massenversand_template.active

source_counter = 2
output_counter = 2

for row in source_sheet:
    output_sheet['A' + str(output_counter)] = source_sheet['A' + str(source_counter)].value       #ID
    output_sheet['B' + str(output_counter)] = source_sheet['B' + str(source_counter + 1)].value   #Event Name
    output_sheet['C' + str(output_counter)] = source_sheet['C' + str(source_counter)].value       #Geschlecht
    output_sheet['D' + str(output_counter)] = source_sheet['D' + str(source_counter)].value       #Name
    output_sheet['E' + str(output_counter)] = source_sheet['E' + str(source_counter)].value       #Vorname
    output_sheet['F' + str(output_counter)] = source_sheet['F' + str(source_counter)].value       #Strasse
    output_sheet['G' + str(output_counter)] = source_sheet['G' + str(source_counter)].value       #Hausnummer
    output_sheet['H' + str(output_counter)] = source_sheet['H' + str(source_counter)].value       #PLZ
    output_sheet['I' + str(output_counter)] = source_sheet['I' + str(source_counter)].value       #Ort
    output_sheet['J' + str(output_counter)] = source_sheet['J' + str(source_counter + 1)].value   #Spike qualitativ
    output_sheet['K' + str(output_counter)] = source_sheet['K' + str(source_counter + 1)].value   #Spike quantitativ
    output_sheet['L' + str(output_counter)] = source_sheet['L' + str(source_counter + 1)].value   #NuC qualitativ

    source_counter += 2
    output_counter += 1

wb_massenversand_template.save('output/massenversand.xlsx')

