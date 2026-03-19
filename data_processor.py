import pandas as pd
import os
import sys

def preprocess_engineering_data(df):
    """
    엑셀 데이터를 받아 변형률 계산, 상태(Loading/Unloading), 
    사이클 번호를 판단하여 데이터프레임에 추가합니다.
    """
    print("-> 데이터를 분석하고 전처리하는 중...")
    
    # 원본 훼손 방지를 위해 복사본 사용
    processed_df = df.copy()

    # 필수 열 존재 여부 확인 (대소문자 구분 주의)
    required_cols = ['Strain', 'Stress']
    if not all(col in processed_df.columns for col in required_cols):
        print(f"에러: 엑셀 파일에 필수 열('{required_cols}')이 없습니다.")
        sys.exit()

    # [1단계] Transfer Strain 만들기 (Strain / 100)
    processed_df['Transfer_Strain'] = processed_df['Strain'] / 100

    # [2단계] State (Loading/Unloading) 판단하기
    # 현재 데이터와 이전 데이터의 차이를 계산 (Strain 증가 여부 확인)
    strain_diff = processed_df['Transfer_Strain'].diff()
    processed_df['State'] = strain_diff.apply(lambda x: 'Loading' if x > 0 else 'Unloading')
    # 첫 번째 행은 이전 데이터가 없으므로 Loading으로 기본 설정
    processed_df.loc[0, 'State'] = 'Loading'

    # [3단계] Cycle (사이클 번호) 계산하기
    # Unloading이었다가 Loading으로 바뀌는 시점을 찾아 카운트
    prev_state = processed_df['State'].shift(1)
    cycle_start_points = ((prev_state == 'Unloading') & (processed_df['State'] == 'Loading')).astype(int)
    # 누적합(cumsum)을 이용해 사이클 번호 매기기 (1부터 시작)
    processed_df['Cycle'] = cycle_start_points.cumsum() + 1

    return processed_df

def extract_target_cycles(df, target_cycles=[5, 10, 15]):
    """
    전처리된 데이터에서 Loading 상태이면서 목표 사이클(5, 10, 15회차)인 
    데이터만 필터링하여 반환합니다.
    """
    print(f"-> {target_cycles}회차 Loading 데이터만 추출하는 중...")
    filtered_df = df[(df['State'] == 'Loading') & (df['Cycle'].isin(target_cycles))]
    return filtered_df

def export_results_to_excel(df, original_filepath, output_filename="분석결과_추출데이터.xlsx"):
    """추출된 데이터를 원본 파일이 있는 폴더에 새 엑셀 파일로 저장합니다."""
    # 원본 파일이 있는 폴더 경로 가져오기
    save_dir = os.path.dirname(original_filepath)
    save_path = os.path.join(save_dir, output_filename)

    # 그래프 그리기에 필요한 핵심 열만 정리해서 저장
    cols_to_export = ['Strain', 'Transfer_Strain', 'Stress', 'Cycle', 'State']
    export_df = df[cols_to_export]
    
    try:
        export_df.to_excel(save_path, index=False)
        print(f"-> [성공] 추출된 데이터가 다음 경로에 저장되었습니다: {save_path}")
    except Exception as e:
        print(f"에러: 결과 파일을 저장하지 못했습니다. {e}")