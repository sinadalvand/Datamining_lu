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

# make new excel to save result matrix
resultFile = Workbook()
resultSheet = resultFile.active

rows = [[sheet.cell_value(row, col) for col in range(columnsCount)] for row in range(rowsCount)]
[[resultSheet.cell(row1 + 1, row2 + 1, math.sqrt(sum([pow(abs(rows[row1][col] - rows[row2][col]), 2) for col in range(columnsCount)]))) for row2 in range(rowsCount)] for row1 in range(rowsCount)]

# save result
resultFile.save('euclideanOutput.xls')
