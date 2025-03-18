import pandas as pd

# 파일 경로를 설정합니다.
file_path = 'PAM_check.xlsx'

# 엑셀 파일을 불러옵니다.
df = pd.read_excel(file_path, sheet_name='Sheet1')

# 패턴을 확인하는 함수 정의
def check_patterns(sequence):
    patterns = ["*TT*", "*AA*", "*CC*", "*GG*"]
    return 'O' if any(pattern in sequence for pattern in patterns) else 'X'

# 새 열 추가
df['Pattern_Check'] = df['Highlighted Sequence'].apply(check_patterns)

# 결과를 확인합니다.
print(df.head())

# 결과를 저장할 경우
output_path = 'PAM_check_results.xlsx'
df.to_excel(output_path, index=False)
print(f"결과가 {output_path}에 저장되었습니다.")
