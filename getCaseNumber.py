import xlrd
import xlwt


def getTotalNumberEachModule():

    Moudle = 'Accounts'
    status = 'Passed'
    Fliename = r'D:\Users\wang.junlin\桌面\testcases.xlsx'
    data = xlrd.open_workbook(Fliename, 'rb')
    sheets = data.sheet_by_index(0)
    columns = sheets.col_slice(2, start_rowx=1)
    PassedNumber = 0
    KnownIssueNumber = 0
    for i in columns:
        i = str(i)
        if "Passed" in i:
            PassedNumber += 1
        elif "Knownissue" in i:
            KnownIssueNumber += 1
        # print(type(str(i)))

    # print(PassedNumber, KnownIssueNumber)
    # print(type(columns))
    return PassedNumber


t = getTotalNumberEachModule()
print(t)
print(type(t))