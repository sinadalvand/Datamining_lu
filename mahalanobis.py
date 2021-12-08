# for importing data from xls file
import xlrd as x

# use to writing new sheet
from openpyxl import Workbook

# use to inverse array
import numpy as np

# read dataset from excel file
exfile = x.open_workbook('dataset.xls', True)
sheet = exfile.sheet_by_name('forestfires')

# get number of rows and columns in dataset
rowsCount = sheet.nrows
columnsCount = sheet.ncols

mainMatrix = [[sheet.cell_value(y, x) for y in range(rowsCount)] for x in range(columnsCount)]
averageMatrix = [sum([mainMatrix[col][row] for row in range(rowsCount)]) for col in range(columnsCount)]
covMatrix = [[(1 / (rowsCount - 1)) * sum([((mainMatrix[col1][row] - averageMatrix[col1]) * (mainMatrix[col2][row] - averageMatrix[col2])) for row in range(rowsCount)]) for col2 in range(columnsCount)] for col1 in range(columnsCount)]
invMatrix = np.linalg.inv(covMatrix)
diffMatrix = [[(averageMatrix[col] - mainMatrix[col][row]) for col in range(columnsCount)] for row in range(rowsCount)]
transposeMatrix = [[diffMatrix[j][i] for j in range(len(diffMatrix))] for i in range(len(diffMatrix[0]))]
mahalaMatrix = [[sum([(transposeMatrix[i][k] * invMatrix[k][j]) for k in range(len(invMatrix))]) for j in range(len(invMatrix[0]))] for i in range(len(transposeMatrix))]

# make new excel to save result matrix
resultFile = Workbook()
resultSheet = resultFile.active


final = [[(resultSheet.cell(row=x + 1, column=y + 1, value=mahalaMatrix[x][y])) for y in range(columnsCount)] for x in range(columnsCount)]
resultFile.save('mahalaOutput.xls')
