import pandas as pd
import requests as r
import json
from datetime import date, timedelta

# Create an empty DataFrame to store all the data
all_data = pd.DataFrame()

# Specify the base URL
base_url = 'https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?t=D&d='

# Specify the number of days to collect data for
num_days = 5

# Get the current date
current_date = date.today()

excel_writer = pd.ExcelWriter('櫃買中心.xlsx')

# Iterate over the past 30 days
for i in range(num_days):
    # Calculate the date for the request
    request_date = current_date - timedelta(days=i)
    
    # Subtract 1911 from the year
    adjusted_year = request_date.year - 1911
    
    # Format the date string as YY/MM/DD
    formatted_date = "{:02d}/{:02d}/{:02d}".format(adjusted_year, request_date.month, request_date.day)
    
    # Construct the complete URL
    url = base_url + formatted_date
    
    # Send GET request and get the response
    response = r.get(url)
    print(url)
    # Parse JSON content
    data = response.json()
    
    # Extract relevant data and create a DataFrame
    df = pd.DataFrame.from_dict(data['aaData'])

    try:
        df = df.iloc[:, [0, 1, 10, 13, 22, 23]]
    except Exception as e:
        print(e)
        continue
    
    df = df.rename(columns={0: '代碼', 1: '名字', 10: '外陸資', 13: '投信',22:'自營商', 23: '三大法人'})
    df = df[df['代碼'].str.len() == 4]
    
    # Append the data to the all_data DataFrame
    all_data = pd.concat([all_data, df], ignore_index=True)
    
    # Save the data to a separate sheet in the Excel file
    sheet_name = request_date.strftime("%Y-%m-%d")
    all_data.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Save and close the Excel writer
excel_writer._save()
