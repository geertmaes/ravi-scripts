import json
from tabulate import tabulate
import xlwt
from tempfile import TemporaryFile

#excel things
book = xlwt.Workbook()
sheet1 = book.add_sheet('Right Algorithms')
name = 'active-configuration.xls'
#define header names
col_names = ['DEFINITION', 'ALGORITHM']
sheet1.write(0, 0, col_names[0])
sheet1.write(0, 1, col_names[1])

resultList = []

with open('../input-json.json', 'r') as active_configuration_file:
    active_configuration = json.load(active_configuration_file)
    configurations = active_configuration['grossConfigurations']
    for configuration in configurations:
        configuredElements = configuration['configuredElements']
        for element in configuredElements:
            resultList.append([element['wageElementDefinition'], element['wageElementCalculationConfiguration']['right']['algorithmIdentifier']])
for index, value in enumerate(resultList, 1):
    sheet1.write(index, 0, value[0])
    sheet1.write(index, 1, value[1])
book.save(name)
book.save(TemporaryFile())
print('The following table has been exported to an Excel file')
print(tabulate(resultList, headers=col_names, tablefmt='grid'))
