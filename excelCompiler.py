import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np
import glob
import os

## get dictionary of all sheets in all workbooks ##
sheetdict = dict()
for xlsxfile in glob.glob("./To be compiled/*.xlsx"):
    wb = load_workbook(xlsxfile)
    sheetdict[xlsxfile.replace('./To be compiled\\','').replace('.xlsx','')] = []
    for i in wb.worksheets:
        if i.sheet_state == "visible":
            sheetdict[xlsxfile.replace('./To be compiled\\','').replace('.xlsx','')].append(i.title)

## dftobecompiled - dataframe to be compiled to the correct format ##
ctemplatepath = 'TemplateTableFormats.xlsx'
cwstypelist = ['Water','Steel','Grass']  # material name list
formattype = 'Academic' # database type
newcollist = ['growth_rate','eng_ger_jap_name'] # must contain all items from mergecoldict
renamedict = {'japanese_name':'jap_name'} # list of columns to rename
mergecoldict = {'eng_ger_jap_name':[[' '], 'name', 'german_name','jap_name']} # merged col dict
# {merged column name : [[merge separator], col values to merge]}
iminstructdict = {'growth_rates_list.xlsx': ['name','growth_rate','name','growth_rate']}
compilermode = 'master' # compiler mode which determines how sheets should be compiled
# standard = 1 sheet per excel file
# master = all sheets in 1 excel file
# polymerise = all sheets in 1 sheet in 1 excel file

## Loading dataframe for current compilation task ##
def getrawdf(xlsxfile, cwstype):
    wb = load_workbook(xlsxfile)
    ws = wb[cwstype]
    ## converting excel ws to dataframe ##
    wsdf = pd.DataFrame(ws.values)
    wsdf.columns = wsdf.iloc[0]
    wsdf = wsdf[1:]
    return wsdf

## Loading the correct format for current compilation task ##
def getcorrectformat(templatesheet, cwstype, formattype):
    correctformat = []
    for worksheetname in templatesheet:
        for row in range(1, templatesheet.max_row + 1):
            if cwstype == templatesheet.cell(row=row, column=1).value:
                for col in range(1, templatesheet.max_column + 1):
                    correctformat.append(templatesheet.cell(row=row, column=col).value)
                correctformat.pop(0)
                while correctformat[-1] is None:
                    del correctformat[-1]
                return correctformat
