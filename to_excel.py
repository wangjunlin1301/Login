import pandas  as pd
import os
from openpyxl.workbook import Workbook
import time
import requests
from datetime import timedelta, datetime
from configparser import ConfigParser

config = ConfigParser()
config.read('config.ini',encoding='utf-8')

yesterday = datetime.today() + timedelta(-1)
yesterday_format = yesterday.strftime('%Y-%m-%d')
Staday = datetime.today() + timedelta(-2)
Staday_format = Staday.strftime('%Y-%m-%d')
Friday = datetime.today() + timedelta(-3)
Friday_format = Friday.strftime('%Y-%m-%d')
today = datetime.today()
today_format = today.strftime('%Y-%m-%d')

todayfile = 'D:\desktop/csv/%sCaseID.csv'%today_format
yesterdayfile = 'D:\desktop/csv/%sCaseID.csv'%yesterday_format
Fridayfile = 'D:\desktop/csv/%sCaseID.csv'%Friday_format
Stadayfile = 'D:\desktop/csv/%sCaseID.csv'%Staday_format

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Content-Length': '182',
    'Content-Type': 'application/x-www-form-urlencoded',
    
    'Cookie': 'ga=GA1.2.1810855833.1582793738; wfx_unq=dJfg4Rn4qpkKpbBR; _gid=GA1.2.1001969726.1583715874; tr_session=e02f4524-71d5-405b-a0e9-371ce010cd07; tr_rememberme=351%3AsR.7g6V19QLtOf.2aIdD-h9QqY%2Fzu6ts5ScPl%2FA0D; ADRUM=s=1583806895521&r=https%3A%2F%2Ftest5.t5.blackline.corp%2Fhome%3F0',
    'Host': 'titx02.blackline.corp',
    'Referer': 'https://titx02.blackline.corp/index.php?/milestones/view/3180',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.88 Safari/537.36'
        
    }
session = requests.session()
requests.urllib3.disable_warnings()

def Csv2Excel(csv_name,Excel_name):
    load = os.getcwd()+'\/'    
    trans = pd.read_csv(load[:-1]+csv_name)
    New = pd.ExcelWriter(load[:-1]+Excel_name)
    trans.to_excel(New,index=False)
    New.save()

def Csv2Excel1():
    filedate = str(time.strftime("%Y-%m-%d"))     
    # trans = pd.read_csv("D:\desktop/2020Feb/%sdailyissue.csv"%filedate,usecols = [11,1,0,4,320,20,159])
    trans = pd.read_csv("D:\desktop/2020May/%sdailyissue.csv"%filedate,usecols = [1,11])
    New = pd.ExcelWriter("D:\desktop/2020May/%sdailyissue.xlsx"%filedate,usecols=[1,11])
    trans.to_excel(New,index=False)
    New.save()

def jira_request(method,url,data=None,info=None):
    if method == "POST":
        response = session.post(url=url,data=data,verify=False,headers=headers)
    elif method == "GET":
        response = session.get(url)
    response.encoding='utf-8'
    return response

def FilterStatus(file,value):
    csv1 = pd.read_csv(file,encoding='gb18030')
    csv2 = csv1[(csv1.Status == value)]
    return csv2

def TodayNewFail():
    try:
        todaydata = FilterStatus(todayfile,'Failed')
        if os.path.exists(yesterdayfile):
            yesterdaydata = FilterStatus(yesterdayfile,'Failed')
        elif os.path.exists(Stadayfile):
            yesterdaydata = FilterStatus(Stadayfile,'Failed')
        elif os.path.exists(Fridayfile):
            yesterdaydata = FilterStatus(Fridayfile,'Failed')
        FinalData = pd.merge(yesterdaydata,todaydata,on=['ID','Case ID','Defects','Status'],how='right',indicator='New')    
        FinalData.New = FinalData.New.cat.set_categories(['Old','New','right_only','both'])    
        j = 0
        for i in FinalData.New:
            if i == 'right_only': 
                FinalData.at[j,'New'] ='New'
            elif i=='both':
                FinalData.at[j,'New'] = 'Old'
            j += 1    
        Excel = pd.ExcelWriter("D:\desktop/2020May/Testcase.xlsx")
        FinalData.to_excel(Excel,index=True)
        Excel.save()
    except:
        print('The compared file is not existed, please confirm it!')

def Testrail():
    #进如Testrail，获取cookies
    TestrailUrl = 'https://titx02.blackline.corp/index.php?/auth/login/'
    Auth = {
        'name':config['user']['username'],
        'password': config['user']['password'],
        'rememberme': '1'
    }
    test1 = jira_request("POST",TestrailUrl,data=Auth)
    print("Testrail Login!")
    #导出文件 以及需要的数据
    data = {
        'columns': 'tests:id,tests:original_case_id,tests:defects,tests:status_id',
        'layout': 'tests',
        'separator_hint': '1',
        'format': 'csv',
        '_token':'waqnLBQnXke.2K818GEM'
    }
    milestonenumber = config['url']['MileStoneNumber']
    ExportCsvUrl = 'https://titx02.blackline.corp/index.php?/milestones/export_csv/%s'%milestonenumber
    print(ExportCsvUrl)
    export = jira_request("POST",ExportCsvUrl,data=data)
   
    print("downloading....")
    print(export.status_code)
 
    #写数据
    with open("D:\desktop/csv/%sCaseID.csv"%today_format,'wb') as f:
        for i in export.iter_content():
            f.write(i)
    # #转换以及清洗数据
    TodayNewFail()
    input("File created!Please hit Enter to quit.")

# Testrail()
# TodayNewFail()


