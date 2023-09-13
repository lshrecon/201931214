import pandas as pd
import shutil
import os

# "sorted_filenames.xlsx" 파일에서 상위 300개의 파일 이름을 불러옵니다.
df = pd.read_excel("sorted_filenames.xlsx", index_col=0)
top_300_files = df['File Name'][:300].values.tolist()

# 파일 이름에 확장자를 추가합니다.
top_300_files = [f"{file}.xlsx" for file in top_300_files]

# 모든 파일들이 위치한 디렉토리
src_directory = r"C:\Users\user\OneDrive\바탕 화면\키움API\ETFcr"

# 파일을 옮길 대상 디렉토리
dest_directory = r"C:\Users\user\OneDrive\바탕 화면\키움API\ETF_unselected"

# 대상 디렉토리가 존재하지 않으면 디렉토리를 생성합니다.
if not os.path.exists(dest_directory):
    os.makedirs(dest_directory)

# 디렉토리에 있는 모든 파일을 확인합니다.
for file_name in os.listdir(src_directory):
    # 파일 이름이 top_300_files에 없는 경우, 파일을 옮깁니다.
    if file_name not in top_300_files:
        src_file = os.path.join(src_directory, file_name)
        dest_file = os.path.join(dest_directory, file_name)
        shutil.move(src_file, dest_file)  # 파일을 옮깁니다.
