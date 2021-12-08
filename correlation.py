# for importing data from xls file
import xlrd as x

# use to writing new sheet
from openpyxl import Workbook

# use to calculate sqrt
import math

# read dataset from excel file
exfile = x.open_workbook('dataset.xls', True)
sheet = exfile.sheet_by_name('forestfires')

# get number of rows and columns in dataset
rowsCount = sheet.nrows
columnsCount = sheet.ncols

# main matrix
mainMatrix = [[sheet.cell_value(y, x) for y in range(rowsCount)] for x in range(columnsCount)]


# correlation calculator
def correlation(x, y):
    subx = [i - (sum(x) / len(x) * 1.0) for i in x]
    suby = [i - (sum(y) / len(y) * 1.0) for i in y]
    num = sum([subx[i] * suby[i] for i in range(len(subx))])
    standardDevX = sum([math.sqrt(subx[i]) for i in range(len(subx))])
    standardDevY = sum([math.sqrt(suby[i]) for i in range(len(suby))])
    res = (standardDevX ** 0.5) * (standardDevY ** 0.5)
    cor = num / res
    return cor


# make new excel to save result matrix
resultFile = Workbook()
resultSheet = resultFile.active

# save each index in excel file
[[(resultSheet.cell(row=i + 1, column=j + 1, value=correlation(mainMatrix[i], mainMatrix[j]))) for j in range(columnsCount)] for i in range(columnsCount)]
resultFile.save('correOutput.xls')
