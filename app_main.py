# 분리해둔 모듈(다른 .py 파일)에서 필요한 함수들만 import 합니다.
from data_loader import get_excel_filepath, load_excel_to_dataframe
from data_processor import (preprocess_engineering_data, 
                            extract_target_cycles, 
                            export_results_to_excel)
from chart_drawer import set_matplotlib_font, draw_all_data_chart, draw_cycle_chart

# 분석 대상 사이클 상수로 정의
TARGET_CYCLES = [5, 10, 15]

def main():
    """프로그램 전체 실행 흐름을 제어하는 메인 함수"""
    print("=== [동일고무벨트] Biaxial Strain 데이터 분석 프로그램 시작 ===")
    
    # 0. 그래프 폰트 세팅
    set_matplotlib_font()

    # 1. 파일 선택 (data_loader 모듈)
    filepath = get_excel_filepath()

    # 2. 데이터 로드 (data_loader 모듈)
    df = load_excel_to_dataframe(filepath)

    # 3. 데이터 분석 및 전처리 (data_processor 모듈)
    processed_df = preprocess_engineering_data(df)

    # 4. 목표 데이터 추출 (data_processor 모듈)
    filtered_df = extract_target_cycles(processed_df, TARGET_CYCLES)

    # 5. 추출된 데이터 엑셀로 저장 (data_processor 모듈)
    export_results_to_excel(filtered_df, filepath)

    # 6. 첫 번째 그래프 그리기 (chart_drawer 모듈)
    draw_all_data_chart(processed_df)

    # 7. 두 번째 그래프 그리기 (chart_drawer 모듈)
    draw_cycle_chart(filtered_df, TARGET_CYCLES)

    print("=== 모든 분석 작업이 성공적으로 완료되었습니다! ===")

if __name__ == "__main__":
    # app_main.py를 직접 실행했을 때만 main() 함수를 호출합니다.
    # 만약 다른 파일에서 이 파일을 import 했다면 main()은 실행되지 않습니다.
    main()