import pandas as pd
import glob
import os

# .xlsx 파일이 있는 디렉토리의 경로를 입력합니다.
path = r'C:\Users\user\OneDrive\바탕 화면\키움API\ETFcr' 
all_files = glob.glob(path + "/*.xlsx")

volume_avg_dict = {}

# 각 파일을 읽어들이고, '거래량' 컬럼의 평균값을 계산합니다.
for filename in all_files:
    df = pd.read_excel(filename)
    volume_avg = df['거래량'].mean()  # '거래량' 컬럼을 변경하여 적용해야 할 수 있습니다.
    volume_avg_dict[filename] = volume_avg

# 평균 거래량이 큰 순서로 파일을 정렬합니다.
sorted_files = sorted(volume_avg_dict.items(), key=lambda x: x[1], reverse=True)

file_names = []
# 정렬된 파일 이름을 출력합니다.
for filename, volume_avg in sorted_files:
    base_name = os.path.basename(filename)  # 파일의 기본 이름을 가져옵니다.
    file_name, ext = os.path.splitext(base_name)  # 확장자를 제거합니다.
    file_names.append(file_name)

# 리스트를 DataFrame으로 변환합니다.
df = pd.DataFrame(file_names, columns=['File Name'])

# DataFrame을 엑셀 파일로 저장합니다.
df.to_excel("sorted_filenames.xlsx")