import yfinance as yf
import pandas as pd
import requests as r
import re
from datetime import datetime, timedelta


end_date = datetime.now().strftime('%Y-%m-%d')

start_date = (datetime.now().date() - timedelta(days=30)).strftime('%Y-%m-%d')



res = r.get("http://isin.twse.com.tw/isin/C_public.jsp?strMode=2")
df = pd.read_html(res.text)[0]
df.columns = ['有價證券代號及名稱', '國際證券辨識號碼(ISIN Code)', '上市日', '市場別', '產業別', 'CFICode', '備註']

stock_symbols = []

for code in df['有價證券代號及名稱']:
    pattern = "[0-9]"
    stackCode = code.split('　')
    if stackCode[0] and re.search(pattern, stackCode[0]) and len(stackCode[0]) == 4:
        stock_symbols.append(str(stackCode[0]))


writer = pd.ExcelWriter('YAHOO.xlsx', engine='xlsxwriter')
all_data = pd.DataFrame()


for stock_symbol in stock_symbols:
    try:

        data = yf.download(stock_symbol+'.TW', start=start_date, end=end_date)

        if not data.empty:

            closing_prices = data['Close']
            volume = data['Volume']

            stock_data = pd.DataFrame({'代码': stock_symbol,
                                       '日期': closing_prices.index.date,
                                       '收盤價': closing_prices,
                                       '成交量': volume})
            
        all_data = pd.concat([all_data, stock_data], ignore_index=True)


    except Exception as e:
        print(f"获取股票 {stock_symbol} 数据时出现错误: {str(e)}")
        

date_groups = all_data.groupby('日期')
for date, group in date_groups:
    sheet_name = date.strftime('%Y-%m-%d')
    group.drop('日期', axis=1).to_excel(writer, sheet_name=sheet_name, index=False)

writer._save()
