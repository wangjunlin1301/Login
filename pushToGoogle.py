#encoding = utf-8

import pygsheets
import pandas as pd
from getCaseNumber import getTotalNumberEachModule

# Google Api认证
googleauth = pygsheets.authorize(
    service_file='./khalil-test-278608-faf5f9854726.json')

sheetName = ['Accounts', 'CIM', 'Journals', 'Match', 'ICH', 'Task', 'Variance']

#open the google spreadsheet ('pysheeetsTest' exists)
sh = googleauth.open('Test for Regression')


# 遍历上传
def Upload():
    for SN in sheetName:
        StatusDict = getTotalNumberEachModule(SN)
        Push(StatusDict)


#select the first sheet
def Push(StatusDict):

    # 打开需要上传数据的相应模块工作表
    Workspace = sh.worksheet_by_title(StatusDict['Module'])

    # case数量赋值
    Passed = StatusDict['Passed']
    Failed = StatusDict['Failed']
    KnownIssue = StatusDict['KnownIssue']
    Blocked = StatusDict['Blocked']
    Untested = StatusDict['Untest']
    Total = StatusDict['Total']
    NumberList = [Passed, Failed, KnownIssue, Blocked, Untested, Total]

    # 创建上传的dataframe
    df = pd.DataFrame(columns=NumberList)
    Workspace.set_dataframe(df, (2, 6))


if __name__ == "__main__":
    Upload()