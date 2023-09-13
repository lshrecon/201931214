import pandas as pd
import os
from pykiwoom.kiwoom import *
import time

os.chdir('C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\ETFcr')

df = pd.read_excel("merge.xlsx")
df['일자'] = pd.to_datetime(df['일자'], format="%Y%m%d")
df = df.set_index('일자')

# 3개월, 6개월, 12개월 영업일 수익률
return_df_3 = df.pct_change(62)
return_df_6 = df.pct_change(126)
return_df_12 = df.pct_change(252)

# 데이터 프레임으로 만들기
momentum_df_3 = pd.DataFrame(return_df_3.loc["2023-05-22"])
momentum_df_6 = pd.DataFrame(return_df_6.loc["2023-05-22"])
momentum_df_12 = pd.DataFrame(return_df_12.loc["2023-05-22"])

momentum_df_3.columns = ["3M_절대_모멘텀"]
momentum_df_6.columns = ["6M_절대_모멘텀"]
momentum_df_12.columns = ["12M_절대_모멘텀"]

# 모멘텀의 순위 구하기
momentum_df_3['순위'] = momentum_df_3['3M_절대_모멘텀'].rank(ascending=False)
momentum_df_6['순위'] = momentum_df_6['6M_절대_모멘텀'].rank(ascending=False)
momentum_df_12['순위'] = momentum_df_12['12M_절대_모멘텀'].rank(ascending=False)

# 정렬하기
momentum_df_3 = momentum_df_3.sort_values(by='순위')
momentum_df_6 = momentum_df_6.sort_values(by='순위')
momentum_df_12 = momentum_df_12.sort_values(by='순위')

# 종목명 추가하기
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)
codes = momentum_df_3.index
names = [kiwoom.GetMasterCodeName(code) for code in codes]

momentum_df_3['종목명'] = pd.Series(data=names, index=momentum_df_3.index)
momentum_df_6['종목명'] = pd.Series(data=names, index=momentum_df_6.index)
momentum_df_12['종목명'] = pd.Series(data=names, index=momentum_df_12.index)

# 각 데이터프레임을 엑셀파일로 저장
momentum_df_3[0:].to_excel("3M_absolute_momentum_list.xlsx")
momentum_df_6[0:].to_excel("6M_absolute_momentum_list.xlsx")
momentum_df_12[0:].to_excel("12M_absolute_momentum_list.xlsx")