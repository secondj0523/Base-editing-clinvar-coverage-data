import pandas as pd
import os

# 파일 경로 확인
file_path = 'highlighted_sequences.xlsx'
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File not found: {file_path}")

# 데이터 로드
data = pd.read_excel(file_path)
if 'Sequence' not in data.columns or 'OriginalBase' not in data.columns:
    raise ValueError("The input file must contain 'Sequence' and 'OriginalBase' columns.")

# 염기서열 강조 표시 함수 정의
def highlight_sequence(row):
    sequence = row['Sequence']  # 염기서열
    original_base = row['OriginalBase']  # 원래 염기
    highlighted_sequence = sequence  # 강조 표시된 염기서열 초기화

    # 강조 표시된 염기의 위치 찾기 ([] 괄호의 시작 인덱스)
    try:
        base_index = sequence.index('[')  # 괄호 시작 위치
        base_position = base_index + 1    # [] 괄호 안의 실제 염기 위치
    except ValueError:
        return sequence  # [] 괄호가 없으면 원본 반환

    # 강조 표시 추가 로직
    if original_base == 'C':
        start_pos = base_position + 5
        end_pos = base_position + 10
        for i in range(start_pos, min(end_pos, len(sequence) - 1)):
            if sequence[i:i+2] == "AA":
                highlighted_sequence = (
                    highlighted_sequence[:i] + "*AA*" + highlighted_sequence[i+2:]
                )
                break

        start_pos = base_position - 13
        end_pos = base_position - 20
        for i in range(max(end_pos, 0), start_pos):
            if sequence[i:i+2] == "CC":
                highlighted_sequence = (
                    highlighted_sequence[:i] + "*CC*" + highlighted_sequence[i+2:]
                )
                break

    elif original_base == 'G':
        start_pos = base_position + 13
        end_pos = base_position + 20
        for i in range(start_pos, min(end_pos, len(sequence) - 1)):
            if sequence[i:i+2] == "GG":
                highlighted_sequence = (
                    highlighted_sequence[:i] + "*GG*" + highlighted_sequence[i+2:]
                )
                break

        start_pos = base_position - 5
        end_pos = base_position - 10
        for i in range(max(end_pos, 0), start_pos):
            if sequence[i:i+2] == "TT":
                highlighted_sequence = (
                    highlighted_sequence[:i] + "*TT*" + highlighted_sequence[i+2:]
                )
                break

    return highlighted_sequence

# 강조 표시 열 추가
data['Highlighted Sequence'] = data.apply(highlight_sequence, axis=1)

# 엑셀 파일 저장
output_path = 'PAM_check.xlsx'
data.to_excel(output_path, index=False)
print(f"Output saved to {output_path}")

