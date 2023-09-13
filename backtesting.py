import pandas as pd
import numpy as np
from pykiwoom.kiwoom import *
import matplotlib.pyplot as plt

# 키움증권 OpenAPI+에 로그인
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

# result_3M_final.xlsx 파일 읽기
result_3M = pd.read_excel("result_3M_final.xlsx")

# 상위 3개 종목의 과거 일봉 데이터를 저장할 딕셔너리
stock_data = {}

# 상위 3개 종목 코드 추출
top_3_stocks = result_3M['종목 코드'].head(6)

# 상위 3개 종목에 대해 과거 일봉 데이터를 가져옵니다.
for code in top_3_stocks:
    # 일봉 데이터를 가져옵니다. 입력은 (종목코드, 조회 시작일, 조회 종료일)입니다.
    df = kiwoom.block_request("opt10081",
                               종목코드=code,
                               기준일자='20230517',
                               수정주가구분=1,
                               output="주식일봉차트조회",
                               next=0)
    # '현재가'를 숫자로 변환합니다.
    df['현재가'] = pd.to_numeric(df['현재가'])
    # 종목별로 데이터를 저장합니다.
    stock_data[code] = df

# 비중 설정
weights = np.array([1/6, 1/6, 1/6, 1/6, 1/6, 1/6])

# 포트폴리오의 일일 수익률 계산
portfolio_daily_returns = np.sum([stock_data[stock]['현재가'].pct_change() * weights[i] for i, stock in enumerate(top_3_stocks)], axis=0)

# 포트폴리오의 누적 수익률 계산
portfolio_cum_returns = (1 + portfolio_daily_returns).cumprod() - 1

# 누적 수익률 그래프 그리기
portfolio_cum_returns.plot()
plt.show()