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
cwstypelist = ['Fire','Flying']  # material name list
formattype = 'Basic' # database type
newcollist = ['growth_arate','eng_ger_jap_name'] # must contain all items from mergecoldict
renamedict = {'japanese_name':'jap_name'} # list of columns to rename
mergecoldict = {'eng_ger_jap_name':[['+'], 'name', 'german_name','jap_name']} # merged col dict
# {merged column name : [[merge separator], col values to merge]}
# indexmatchlist = [{}, {}]
# enttype = ['Composite','Performance'] # entity type
compilermode = 'polymerise' # compiler mode which determines how sheets should be compiled
# standard = 1 sheet per excel file
# master = all sheets in 1 excel file
# polymerise = all sheets in 1 sheet in 1 excel file