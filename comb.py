import pandas as pd


file1 = pd.read_excel('櫃買中心.xlsx', sheet_name=None)


file2 = pd.read_excel('證券所.xlsx', sheet_name=None)


file3 = pd.read_excel('YAHOO.xlsx', sheet_name=None)


merged_data = {}

for sheet in file1:
    if sheet in merged_data:
        merged_data[sheet] = pd.concat([merged_data[sheet], file1[sheet]], ignore_index=True)
    else:
        merged_data[sheet] = file1[sheet]

for sheet in file2:
    if sheet in merged_data:
        merged_data[sheet] = pd.concat([merged_data[sheet], file2[sheet]], ignore_index=True)
    else:
        merged_data[sheet] = file2[sheet]
for sheet in file3:
    if sheet in merged_data:
        if '代碼' in file3[sheet].columns:
            merged_data[sheet] = merged_data[sheet].merge(file3[sheet], on='代碼', how='left')
        else:
            merged_data[sheet] = pd.concat([merged_data[sheet], file3[sheet]], ignore_index=True)
    else:
        merged_data[sheet] = file3[sheet]
with pd.ExcelWriter('天妮超可愛!.xlsx') as writer:
    for sheet, data in merged_data.items():
        data.to_excel(writer, sheet_name=sheet, index=False)
        