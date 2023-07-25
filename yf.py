import yfinance as yf
import pandas as pd
import requests as r
import re
from datetime import datetime, timedelta
import time
import sys


default_value=30
try:
    inputday = int(sys.argv[1])
except:
    inputday=default_value

end_date = datetime.now().strftime('%Y-%m-%d')

start_date = (datetime.now().date() - timedelta(days=inputday)).strftime('%Y-%m-%d')


data = pd.read_excel('證券所.xlsx', skiprows=range(1), usecols=[1])
data = data.astype(str)

stock_symbols = set(data.values.flatten().tolist())

writer = pd.ExcelWriter('YAHOO.xlsx', engine='xlsxwriter')
all_data = pd.DataFrame()


for stock_symbol in stock_symbols:
    try:

        data = yf.download(stock_symbol+'.TW', start=start_date, end=end_date)

        if not data.empty:

            closing_prices = data['Close']
            volume = data['Volume']

            stock_data = pd.DataFrame({'代碼': stock_symbol,
                                       '日期': closing_prices.index.date,
                                       '收盤價': closing_prices,
                                       '成交量': volume})
            
        all_data = pd.concat([all_data, stock_data], ignore_index=True)


    except Exception as e:
        print(f"{stock_symbol} : {str(e)}")
  
base_url = 'https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?d='



# Iterate over the past 30 days
for i in range(inputday):
    # Calculate the date for the request
    request_date = datetime.now().date() - timedelta(days=i)
    
    # Subtract 1911 from the year
    adjusted_year = request_date.year - 1911
    
    # Format the date string as YY/MM/DD
    formatted_date = "{:02d}/{:02d}/{:02d}".format(adjusted_year, request_date.month, request_date.day)
    
    # Construct the complete URL
    url = base_url + formatted_date
    
    retry_count = 0
    response = None
    
    while retry_count < 3 and response is None:
        try:
            response = r.get(url, timeout=10)
        except r.exceptions.RequestException:
            print("No response. Retrying after {} seconds...".format(1))
            time.sleep(1)
            retry_count += 1
    
    if response is None:
        print("No response after multiple retries. Skipping this request.")
        continue
    # Parse JSON content
    data = response.json()
    
    # Extract relevant data and create a DataFrame
    df = pd.DataFrame.from_dict(data['aaData'])

    try:
        df = df.iloc[:, [0, 2, 8]]
    except Exception as e:
        print(e)
        continue
    
    df = df.rename(columns={0: '代碼', 2: '收盤價', 8: '成交量'})
    df = df[df['代碼'].str.len() == 4]
    df['日期'] = request_date
    
    all_data = pd.concat([all_data, df], ignore_index=True)
    
        

date_groups = all_data.groupby('日期')
unique_data = []

for date, group in date_groups:
    group = group.drop_duplicates()  # Remove duplicate entries for the same date
    if not group.empty:
        unique_data.append(group)

for data_group in unique_data:
    date = data_group['日期'].iloc[0]
    sheet_name = date.strftime('%Y-%m-%d')
    data_group.drop('日期', axis=1).to_excel(writer, sheet_name=sheet_name, index=False)

writer._save()
