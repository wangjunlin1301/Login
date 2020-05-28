#encoding = utf-8

import pygsheets
import pandas as pd

googleauth = pygsheets.authorize(
    service_file='./khalil-test-278608-faf5f9854726.json')

df = pd.DataFrame()
sheetName = ['Accounts', 'CIM', 'Journals', 'Match']

df['Name'] = ['khalil', 'tesasdasdast']

#open the google spreadsheet ('pysheeetsTest' exists)
sh = googleauth.open('Test for Regression')

#select the first sheet
for Name in sheetName:
    wks = sh.worksheet_by_title(Name)

    #update the first sheet with df, starting at cell B2
    wks.set_dataframe(df, (1, 3))
