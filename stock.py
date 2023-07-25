import pandas as pd
import requests as r
from datetime import datetime, timedelta
import sys

default_value=30
try:
    inputday = int(sys.argv[1])
except:
    inputday=default_value

end_date = datetime.now().date()
start_date = end_date - timedelta(days=inputday)
date_range = pd.date_range(start=start_date, end=end_date, freq='D')[::-1]

# 创建一个空的Excel写入对象
writer = pd.ExcelWriter('證券所.xlsx', engine='xlsxwriter')

# 循环获取每一天的数据并写入不同的sheet
for date in date_range:
    time = date.strftime('%Y%m%d')
    url = f'https://www.twse.com.tw/fund/T86?response=json&date={time}&selectType=ALL'
    try:
        res = r.get(url)
        inv_json = res.json()
        df = pd.DataFrame.from_dict(inv_json['data'])
    except Exception as e:
        print(e)
        continue
    
    df = df.iloc[:, [0, 1, 4, 7, 10, 14, 17, 18]]
    df = df.rename(columns={0: '代碼', 1: '名字',4:'外陸資', 7:'外資自營商', 10:'投信', 14:'自營商買賣(自行)', 17:'自營商買賣(避險)', 18: '三大法人'})
    df = df[df['代碼'].str.len() == 4]
    df.insert(0, '日期', date)
    
    # 写入不同的sheet，使用日期作为sheet名
    df.to_excel(writer, sheet_name=date.strftime('%Y-%m-%d'), index=False)

# 保存Excel文件
writer._save()