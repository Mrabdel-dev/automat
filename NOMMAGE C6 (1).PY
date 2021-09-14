# -*- coding: utf-8 -*-
"""
Created on Tue Feb 23 10:00:02 2021

@author: BE24
"""

import openpyxl
from openpyxl import load_workbook

annexeC6 = load_workbook("Annexe C6.xlsx")
export = annexeC6['Export 1']
maxRow = export.max_row
photo = annexeC6['Photos']

ensseCode = export.cell(row=3, column=7).value
print(ensseCode)
lin = 3
for i in range(9, maxRow + 1):
    val = export.cell(row=i, column=1).value
    if val is not None and len(val) < 8:
        photo.cell(row=lin, column=1).value = str(ensseCode) + "_" + val + "_" + "1"
        photo.cell(row=lin, column=2).value = str(ensseCode) + "_" + val + "_" + "2"
        lin += 2

annexeC6.save("anexC6New.xlsx")