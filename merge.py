import pandas as pd
from datetime import timedelta, datetime
from openpyxl.workbook import Workbook
from configparser import ConfigParser

config = ConfigParser()

config.read('config.ini', encoding='utf-8')
print(type(config['user']['username']))
aa = config['filter']['ExportJirabug']
print(aa)
print(type(aa))