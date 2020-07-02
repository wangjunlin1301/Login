import xlrd
import xlwt
from jira import Excelpath,Regression

Fliename = r'../CountCase.xlsx'
File = Excelpath + "/%s.xlsx" % Regression

def getTotalNumberEachModule(Module):

    # 获取文件数据
    data = xlrd.open_workbook(Fliename, 'rb')
    sheets = data.sheet_by_index(0)

    # 模块
    Moudlecol = sheets.col_slice(0, start_rowx=1)
    Rowindex = 1
    Pcols = sheets.col_slice(1, start_rowx=Rowindex)
    StatusCol = sheets.col_slice(2, start_rowx=Rowindex)

    # 定义状态码的案例数量
    PassedNumber = 0
    KnownIssueNumber = 0
    FailedNumber = 0
    BlockedNumber = 0
    UntestNumber = 0
    P0, P1, P2, P3 = 0, 0, 0, 0

    # 筛选模块
    for m in Moudlecol:
        Rowindex += 1
        m = str(m)
        if Module in m:
            i = str(StatusCol[Rowindex - 2])
            # 对字段条件进行判断
            if "Passed" in i:
                PassedNumber += 1
            elif "Known" in i:
                KnownIssueNumber += 1
            elif "Failed" in i:
                FailedNumber += 1
            elif "Blocked" in i:
                BlockedNumber += 1
            elif "Untested" in i:
                UntestNumber += 1
            elif "Retest" in i:
                UntestNumber +=1

            # 对级别条件进行判断
            p = str(Pcols[Rowindex - 2])
            if "Critical" in p:
                P0 += 1
            elif "High" in p:
                P1 += 1
            elif "Medium" in p:
                P2 += 1
            elif "Low" in p:
                P3 += 1

    # 添加值进入字典
    Status = {
        "Module":
        Module,
        "Passed":
        PassedNumber,
        "Failed":
        FailedNumber,
        "KnownIssue":
        KnownIssueNumber,
        "Blocked":
        BlockedNumber,
        "Untest":
        UntestNumber,
        "Total":
        PassedNumber + FailedNumber + KnownIssueNumber + BlockedNumber +
        UntestNumber
    }
    Priority = {'P0': P0, 'P1': P1, 'P2': P2, 'P3': P3}

    return Status, Priority

def getTicketnumber():

    data = xlrd.open_workbook(File, 'rb')
    sheets = data.sheet_by_index(0)

    # 模块
    Rowindex = 1
    Versioncols = sheets.col_slice(7, start_rowx=Rowindex)
    StatusCol = sheets.col_slice(4, start_rowx=Rowindex)

    newNumber,readyDevNumber,inDevNumber,inQaNumber,closedNumber,PRNumber,PANumber =0,0,0,0,0,0,0
    CnewNumber,CreadyDevNumber,CinDevNumber,CinQaNumber,CclosedNumber,CPRNumber ,CPANumber=0,0,0,0,0,0,0

    for index,i in enumerate([j for j in Versioncols]):
        h = str(StatusCol[index])
        if 'empty' in str(i):
            if 'New' in h:
                newNumber +=1
            elif 'Closed' in h:
                closedNumber +=1
            elif 'In Dev' in h:
                inDevNumber +=1
            elif 'In QA' in h:
                inQaNumber +=1
            elif 'Product Reviewed' in h:
                PRNumber +=1
            elif 'Ready for Dev' in h:
                readyDevNumber +=1
            elif 'Product Accepted' in h:
                PANumber +=1
        else:
            if 'New' in h:
                CnewNumber +=1
            elif 'Closed' in h:
                CclosedNumber +=1
            elif 'In Dev' in h:
                CinDevNumber +=1
            elif 'In QA' in h:
                CinQaNumber +=1
            elif 'Product Reviewed' in h:
                CPRNumber +=1
            elif 'Ready for Dev' in h:
                CreadyDevNumber +=1
            elif  'In Product Acceptance' in h:
                CPANumber +=1

    CurrentFixVersion = {
        'New':CnewNumber,
        'PR':CPRNumber,
        'ReadyDev':CreadyDevNumber,
        'InDev':CinDevNumber,
        'InQA':CinQaNumber,
        'PA':CPANumber,
        'Closed':CclosedNumber,
    }
    NoFixVersion = {
        'New':newNumber,
        'PR':PRNumber,
        'ReadyDev':readyDevNumber,
        'InDev':inDevNumber,
        'InQA':inQaNumber,
        'PA':PANumber,
        'Closed':closedNumber
    }
    return CurrentFixVersion,NoFixVersion