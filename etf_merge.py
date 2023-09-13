import pandas as pd
import os

os.chdir('C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\ETFcr')
flist = os.listdir()
xlsx_list = [x for x in flist if x.endswith('.xlsx')]
close_data = []

for xls in xlsx_list:
    code = xls.split('.')[0]
    df = pd.read_excel(xls)
    df2 = df[['일자', '현재가',]].copy()
    df2.rename(columns={'현재가': code}, inplace=True)
    df2 = df2.set_index('일자')
    df2 = df2[::-1]
    close_data.append(df2)

# concat
df = pd.concat(close_data, axis=1)
df.to_excel("merge.xlsx")