from pykiwoom.kiwoom import *
import datetime
import time

# 로그인
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

# 전종목 종목코드
etf = kiwoom.GetCodeListByMarket('8')
print(len(etf), etf)

# 문자열로 오늘 날짜 얻기
now = datetime.datetime.now()
today = now.strftime("%Y%m%d")

# 저장할 경로
output_dir = 'C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\ETFcr'

# 전 종목의 600일 가격 데이터
for i, code in enumerate(etf):
    print(f"{i}/{len(etf)} {code}")
    df = kiwoom.block_request("opt10081",
                              종목코드=code,
                              기준일자=today,
                              수정주가구분=1,
                              output="주식일봉차트조회",
                              next=0)
    out_name = f"{output_dir}\\{code}.xlsx"
    df.to_excel(out_name)
    time.sleep(3.6)