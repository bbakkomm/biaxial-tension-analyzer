import matplotlib.pyplot as plt

def set_matplotlib_font():
    """Matplotlib에서 한글이 깨지지 않도록 맑은 고딕 폰트를 설정합니다."""
    plt.rcParams['font.family'] = 'Malgun Gothic' # 윈도우 기준
    plt.rcParams['axes.unicode_minus'] = False # 마이너스 기호 깨짐 방지

def draw_all_data_chart(df):
    """원본 데이터 전체를 그리는 그래프 (캡처-1 스타일)"""
    print("-> 전체 데이터 그래프를 생성하는 중...")
    
    # 새로운 그래프 창 생성
    plt.figure(figsize=(10, 6))
    
    # 얇은 실선으로 모든 데이터 플롯
    plt.plot(df['Transfer_Strain'], df['Stress'], color='royalblue', linewidth=1)

    # 제목 및 축 이름 설정
    plt.title('동일고무벨트_Biaxial Tension (전체 데이터)')
    plt.xlabel('Nominal Strain')
    plt.ylabel('Nominal Stress (MPa)')
    
    # 그리드 추가
    plt.grid(True, linestyle='--', alpha=0.7)
    
    # 그래프 보여주기 (팝업창)
    plt.show()

def draw_cycle_chart(df, target_cycles=[5, 10, 15]):
    """특정 사이클만 그리는 그래프 (캡처-2 스타일)"""
    print(f"-> {target_cycles}회차 Loading 그래프를 생성하는 중...")
    
    plt.figure(figsize=(10, 6))

    # 데이터가 끊기지 않고 하나로 이어지는 것을 방지하기 위해 
    # 사이클 번호별로 반복문을 돌려 플롯합니다.
    for cycle in target_cycles:
        cycle_data = df[df['Cycle'] == cycle]
        
        # 선을 조금 더 굵게 표현
        plt.plot(
          cycle_data['Transfer_Strain'], 
          cycle_data['Stress'], 
          color='royalblue', 
          linewidth=1.5, 
          label=f'Cycle {cycle}'
          )

    plt.title(f'동일고무벨트_Biaxial ({", ".join(map(str, target_cycles))}th Loading)')
    plt.xlabel('Engineering Strain')
    plt.ylabel('Engineering Stress (MPa)')
    plt.grid(True, linestyle='--', alpha=0.7)
    
    # 어떤 선이 어떤 사이클인지 보여주는 범례 추가 (데이터가 겹쳐서 안 보일 수도 있음)
    # plt.legend() 
    
    plt.show()