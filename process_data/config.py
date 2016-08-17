import os
import sqlite3
from datetime import date

DIRS = {
    'SOURCE_DIR': 'C:\\Users\\WY-ZY\\OneDrive\\Data\\',
    'BASE_DIR': os.path.dirname(os.path.abspath(__file__)),
}

WORKBOOKS = {
    '产品信息': '产品信息.xlsx',
    '拉拔转序': '2016-8拉拔转序.xlsx',
    '拉拔日志': '拉拔日志.xlsx',
    '一部转序': '2016-8生产一部转序.xlsx',
    '一部工序监控': '2016-8工序监控.xlsm',
    'G加转序': '2016-8月四号厂房转序.xlsx',
    'G加工序数据': '16年4号厂房生产报表.xlsm',
    '东海转序': '8月新各段转序汇总表.xls',
    '库存1': '产成品入库盘点表（2016年7月份）.xlsx',
    '库存2': '副本活塞杆7月份盘点表.xlsx',
    '外协': '8月外协朝阳.xlsx',
}

DATABASES = {
    'default': os.path.join(os.path.dirname(DIRS['BASE_DIR']), 'db.sqlite3')
    }

CONN = sqlite3.connect(database=DATABASES['default'])

# start date
SD = date(2016, 7, 31)
