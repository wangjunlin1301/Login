import pandas as pd
from datetime import timedelta, datetime
from openpyxl.workbook import Workbook
from configparser import ConfigParser

config = ConfigParser()

config.read('config.ini',encoding='utf-8')
print(config['user']['username'])
data = {
    'user':'wang.junlin'
}
print(data['user'])