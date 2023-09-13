import pandas as pd

result_3M = pd.read_excel("C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\result_3M.xlsx")
result_6M = pd.read_excel("C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\result_6M.xlsx")
result_12M = pd.read_excel("C:\\Users\\user\\OneDrive\\바탕 화면\\키움API\\result_12M.xlsx")

# '절대 모멘텀' 컬럼이 비어있는 행 삭제
result_3M = result_3M.dropna(subset=['3M_절대_모멘텀'])
result_6M = result_6M.dropna(subset=['6M_절대_모멘텀'])
result_12M = result_12M.dropna(subset=['12M_절대_모멘텀'])

# '순위' 컬럼을 기준으로 오름차순 정렬
result_3M = result_3M.sort_values(by='순위', ascending=True)
result_6M = result_6M.sort_values(by='순위', ascending=True)
result_12M = result_12M.sort_values(by='순위', ascending=True)

# 컬럼 이름 변경
result_3M = result_3M.rename(columns={"Unnamed: 0": "종목 코드"})
result_6M = result_6M.rename(columns={"Unnamed: 0": "종목 코드"})
result_12M = result_12M.rename(columns={"Unnamed: 0": "종목 코드"})


print(result_3M.head())
print(result_6M.head())
print(result_12M.head())

# 엑셀로 저장
result_3M.to_excel("result_3M_final.xlsx", index=False)
result_6M.to_excel("result_6M_final.xlsx", index=False)
result_12M.to_excel("result_12M_final.xlsx", index=False)