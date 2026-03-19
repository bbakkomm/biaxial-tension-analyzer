import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import sys

def get_excel_filepath():
    """tkinter GUI를 사용해 사용자에게 엑셀 파일 선택 창을 띄웁니다."""
    # tkinter 초기화 (GUI 창은 숨김)
    root = tk.Tk()
    root.withdraw()
    
    # .exe로 빌드했을 때 파일 선택 창이 맨 위로 오게 처리
    root.attributes('-topmost', True) 

    # 파일 선택창 열기
    filepath = filedialog.askopenfilename(
        title="분석할 엑셀 파일을 선택하세요", 
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    # 사용자가 취소를 눌렀을 때 처리
    if not filepath:
        print("에러: 파일 선택이 취소되었습니다. 프로그램을 종료합니다.")
        sys.exit() # 프로그램 강제 종료

    return filepath

def load_excel_to_dataframe(filepath):
    """지정된 경로의 엑셀 파일을 Pandas 데이터프레임으로 로드합니다."""
    print(f"-> [{os.path.basename(filepath)}] 데이터를 불러오는 중...")
    try:
        df = pd.read_excel(filepath)
        print("-> 데이터 로드 완료.")
        return df
    except Exception as e:
        print(f"에러: 엑셀 파일을 열 수 없습니다. {e}")
        sys.exit()