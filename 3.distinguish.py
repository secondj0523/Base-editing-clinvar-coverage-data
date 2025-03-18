import pandas as pd

# 파일 읽기
file_path_final = 'PAM_check_results.xlsx'
data_checked = pd.read_excel(file_path_final)

# 'Highlight Check' 열 생성 (필요시 생성)
if 'Highlight Check' not in data_checked.columns:
    data_checked['Highlight Check'] = data_checked['Highlighted Sequence'].apply(
        lambda x: 'O' if isinstance(x, str) and ('*TT*' in x or '*AA*' in x) else 'X'
    )

# 새 열 추가
data_checked['Contains *TT* or *AA*'] = data_checked['Highlighted Sequence'].apply(
    lambda x: 1 if isinstance(x, str) and ('*TT*' in x or '*AA*' in x) else 0
)
data_checked['Contains *GG* or *CC*'] = data_checked['Highlighted Sequence'].apply(
    lambda x: 1 if isinstance(x, str) and ('*GG*' in x or '*CC*' in x) else 0
)
data_checked['Contains *TT* and *GG*'] = data_checked['Highlighted Sequence'].apply(
    lambda x: 1 if isinstance(x, str) and ('*TT*' in x and '*GG*' in x) else 0
)
data_checked['Contains *AA* and *CC*'] = data_checked['Highlighted Sequence'].apply(
    lambda x: 1 if isinstance(x, str) and ('*AA*' in x and '*CC*' in x) else 0
)
data_checked['O Check Count'] = data_checked['Highlight Check'].apply(
    lambda x: 1 if x == 'O' else 0
)

# 결과를 새로운 엑셀 파일로 저장
output_path_final = 'ABE.xlsx'
data_checked.to_excel(output_path_final, index=False)

print(f"Output saved to {output_path_final}")
