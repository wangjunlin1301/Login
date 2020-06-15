# -*- coding:utf-8 -*-
import requests
import time
import pandas as pd
# from login import to_excel
from configparser import ConfigParser
from to_excel import Comparecases, ExportTestcases

config = ConfigParser()
config.read(r'config.ini', encoding='utf-8')
headers = {
    'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0'
}
session = requests.session()
session.keep_alive = False
requests.urllib3.disable_warnings()

currentdate = str(time.strftime("%Y-%m-%d"))
filedate = str(time.strftime("%Y-%m-%d"))
Csvpath = config['Newpath']['csvpath']
Excelpath = config['Newpath']['excelpath']
Regression = config['filter']['RegressionName']
JiraQuery = config['filter']['ExportJirabug']
JiraQueryAll = config['filter']['ExportJirabugAll']
print(currentdate)


def jira_request(method, url, data=None, info=None):
    if method == "POST":
        response = session.post(url=url,
                                data=data,
                                verify=False,
                                headers=headers)
    elif method == "GET":
        response = session.get(url, headers=headers)
    response.encoding = 'utf-8'
    response.cookies
    return response


def saveXlsxOfBug():
    getBugCsvFile()
    try:
        #回归开始以来的bug
        trans1 = pd.read_csv(Csvpath + "/%s.csv" % Regression,
                             usecols=[1, 3, 5, 6, 13, 12, 9, 14])
        New1 = pd.ExcelWriter(Excelpath + "/%s.xlsx" % Regression)
        trans1.to_excel(New1, index=True)
        New1.save()
        print("Updated daily bug!")
    except:
        print('SaveXlsxOfBug Wrong!')
    try:
        #每天的bug 存储
        trans = pd.read_csv(Csvpath + "/%sdailyissue.csv" % filedate,
                            usecols=[1, 3, 5, 6, 13, 12, 9, 14])
        New = pd.ExcelWriter(Excelpath + "/%sdailyissue.xlsx" % filedate)
        trans.to_excel(New, index=True)
        New.save()

    except:
        print("No new bug!")
    print("文件保存成功！请继续执行pushToGoogle.py！")


def getBugCsvFile():
    LoginJiraUrl = 'https://jira.blackline.corp/login.jsp'
    LoginData = {
        'os_username': config['user']['username'],
        'os_password': config['user']['password']
    }

    t = jira_request('POST', LoginJiraUrl, data=LoginData)
    if t.status_code == '200':
        print('Login Success!')

    # 获取
    GetFileUrl = "https://jira.blackline.corp/sr/jira.issueviews:searchrequest-csv-current-fields/temp/SearchRequest.csv?jqlQuery= created >= " + currentdate + JiraQuery
    GetbugUrlutil = "https://jira.blackline.corp/sr/jira.issueviews:searchrequest-csv-current-fields/temp/SearchRequest.csv?jqlQuery=" + JiraQueryAll
    #GetbugUrlutil = "https://jira.blackline.corp/sr/jira.issueviews:searchrequest-csv-current-fields/temp/SearchRequest.csv?jqlQuery=+(+labels+=+7.26Regression+OR+'Found in Build'+~+'7.26*'+)+AND+issuetype+in+(+Bug+,+'Internal Bug'+)+AND+created+>=+2020-03-08+AND+status+not+in+(Closed)+AND+labels+not+in+(product.not.for.7.26)"
    result = jira_request("GET", url=GetFileUrl)
    result1 = jira_request("GET", url=GetbugUrlutil)
    with open(Csvpath + '/%sdailyissue.csv' % filedate, 'wb') as f:
        for i in result.iter_content():
            f.write(i)
    with open(Csvpath + '/%s.csv' % Regression, 'wb') as f:
        for i in result1.iter_content():
            f.write(i)
    print('Daliy Issue Updating!')
    # 筛选出新失败的cases
    Comparecases()
    # 导出cases
    ExportTestcases()


if __name__ == "__main__":
    saveXlsxOfBug()
