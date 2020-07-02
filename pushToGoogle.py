#encoding = utf-8

import pygsheets
import pandas as pd
from getCaseNumber import getTotalNumberEachModule,getTicketnumber
from configparser import ConfigParser
from to_excel import ExportTestcases
import time
from jira import filedate,Regression,Excelpath

config = ConfigParser()
config.read('config.ini', encoding='utf-8')
GoogleName = config['Newpath']['ExcelName']

# Google Api认证
googleauth = pygsheets.authorize(
    service_file='./khalil-test-278608-faf5f9854726.json')

sheetName = [
    'Accounts', 'CIM', 'Journals', 'Matching', 'Intercompany', 'Tasks',
    'Variance', 'Compliance', 'RAD', 'RFC', 'StarterDB', 'System', 'Users',
    'Daily Reconciliations', 'BLJ'
]

# DaliyIssues 文件
DaliyIssueFile = Excelpath + "/%s.xlsx" % Regression
NewIssueFile = Excelpath + "/%sdailyissue.xlsx" % filedate
#open the google spreadsheet ('pysheeetsTest' exists)
sh = googleauth.open(GoogleName)


# 遍历上传
def Upload():
    for SN in sheetName:
        Dict = getTotalNumberEachModule(SN)
        StatusDict = dict(Dict[0])
        PriorityDict = dict(Dict[1])
        Push(StatusDict, PriorityDict)


def updateGooglesheet():
    try:
        Dictlist = getTicketnumber()
        Workspace = sh.worksheet_by_title('Tables')
        Currentfixdf = pd.DataFrame(columns = [v for k,v in Dictlist[0].items()])
        Nofixversiondf = pd.DataFrame(columns = [i for j,i in Dictlist[1].items()])
        Workspace.set_dataframe(Currentfixdf, (117, 3))
        Workspace.set_dataframe(Nofixversiondf, (118, 3))
    except:
        print('something wrong, please contact author!')
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


def PushTestcases():
    try:
        Testcasesdf = pd.read_excel(Excelpath + '/Testcase.xlsx', skiprows=1)
        Testcasesdf.style.set_properties(**{'text-align': 'left'})
        Workspace = sh.worksheet_by_title('FailedCase')
        Workspace.set_dataframe(Testcasesdf, (2, 1))
        print("Test Cases更新成功！")
    except:
        print("something wrong in test cases file")


def PushDaliyIssues():
    try:
        DaliyIssuedf = pd.read_excel(DaliyIssueFile, index_col=0, skiprows=1)
        colsList = DaliyIssuedf.columns.values.tolist()
        DaliyIssuedf = DaliyIssuedf[[
            colsList[0], colsList[1], colsList[2], colsList[3], colsList[7],
            colsList[5], colsList[4], colsList[6]
        ]]
        DaliyIssuedf.style.set_properties(**{'text-align': 'left'})
        Workspace = sh.worksheet_by_title('DailyIssues')

        # new issue
        newIssueDf = pd.read_excel(NewIssueFile,index_col=0,skiprows=1)
        colsList1 = newIssueDf.columns.values.tolist()
        newIssueDf = newIssueDf[[
                    colsList1[0],colsList1[1],colsList1[2],colsList1[3],colsList1[5],colsList1[4]
            ]]
        newIssueDf.style.set_properties(**{'text-align':'left'})
        Workspace.set_dataframe(newIssueDf,(2,14))
        Workspace.set_dataframe(DaliyIssuedf, (2, 2))
        print("Push Done")
    except:
        print('表格错误！')


if __name__ == "__main__":
    updateGooglesheet()    
    # 上传日常bug
    PushDaliyIssues()
    PushTestcases()
    # # 更新到GoogleSheets
    Upload()
