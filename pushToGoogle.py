#encoding = utf-8

import pygsheets
import pandas as pd
from getCaseNumber import getTotalNumberEachModule
from configparser import ConfigParser
from to_excel import ExportTestcases
import time

config = ConfigParser()
config.read('config.ini', encoding='utf-8')
GoogleName = config['Newpath']['ExcelName']
Excelpath = config['Newpath']['excelpath']
filedate = str(time.strftime("%Y-%m-%d"))

# Google Api认证
googleauth = pygsheets.authorize(
    service_file='./khalil-test-278608-faf5f9854726.json')

sheetName = [
    'Accounts', 'CIM', 'Journals', 'Matching', 'Intercompany', 'Tasks',
    'Variance', 'Compliance', 'RAD', 'RFC', 'StarterDB', 'System', 'Users',
    'Daily Reconciliations', 'BLJ'
]

# DaliyIssues 文件
DaliyIssueFile = Excelpath + "/%sdailyissue.xlsx" % filedate
#open the google spreadsheet ('pysheeetsTest' exists)
sh = googleauth.open(GoogleName)


# 遍历上传
def Upload():
    for SN in sheetName:
        Dict = getTotalNumberEachModule(SN)
        StatusDict = dict(Dict[0])
        PriorityDict = dict(Dict[1])
        Push(StatusDict, PriorityDict)


#select the first sheet
def Push(StatusDict, PriorityDict):

    # 打开需要上传数据的相应模块工作表
    try:
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

        P0, P1, P2, P3 = PriorityDict['P0'], PriorityDict['P1'], PriorityDict[
            'P2'], PriorityDict['P3'],
        PnumberList = [P0, P1, P2, P3]
        pf = pd.DataFrame(columns=PnumberList)
        Workspace.set_dataframe(pf, (2, 14))
        print(StatusDict['Module'] + '  Done!')
    except:
        print(StatusDict['Module'] + ' is not exised!')


def PushDaliyIssues():
    DaliyIssuedf = pd.read_excel(DaliyIssueFile, index_col=0, skiprows=1)
    colsList = DaliyIssuedf.columns.values.tolist()
    DaliyIssuedf = DaliyIssuedf[[
        colsList[0], colsList[1], colsList[2], colsList[3], colsList[7],
        colsList[5], colsList[6], colsList[4]
    ]]
    try:
        Workspace = sh.worksheet_by_title('DaliyIssues')
        Workspace.set_dataframe(DaliyIssuedf, (2, 2))
        print("aaa")
    except:
        print('表格错误！')


if __name__ == "__main__":
    # 上传日常bug
    PushDaliyIssues()
    # 更新到GoogleSheets
    Upload()
