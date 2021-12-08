# for importing data from xls file
from collections import Counter

import xlrd as x

# use to writing new sheet
from openpyxl import Workbook

# use to calculate sqrt , e ,log
from math import e, log, sqrt

# read dataset from excel file
exfile = x.open_workbook('dataset.xls', True)
sheet = exfile.sheet_by_name('forestfires')

# get number of rows and columns in dataset
rowsCount = sheet.nrows
columnsCount = sheet.ncols

#  main Data
MainMatrix = [[sheet.cell_value(row, col) for row in range(rowsCount)] for col in range(columnsCount)]


# calculate entropy for each column
def entropy(col):
    # set count of repetition of data
    counts = Counter()
    for d in col:
        counts[d] += 1
    probability = [float(c) / len(col) for c in counts.values()]
    ent = 0
    for p in probability:
        if p > 0.:
            ent -= p * log(p, e)
    return ent


# make new excel to save result matrix
resultFile = Workbook()
resultSheet = resultFile.active

# save result in excel file
[(resultSheet.cell(row=1, column=x + 1, value=entropy(MainMatrix[x]))) for x in range(len(MainMatrix))]
resultFile.save('entropyOutput.xls')
