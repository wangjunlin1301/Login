import xlrd
import xlwt

Fliename = r'../CountCase.xlsx'


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
