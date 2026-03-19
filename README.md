# 📈 고무벨트 인장시험 데이터 분석기 (Biaxial Tension Analyzer)

고무벨트의 Biaxial Strain(이축 인장) 데이터를 불러와 자동으로 사이클(Cycle)을 분석하고, 사용자가 원하는 구간의 데이터를 추출하여 시각화 및 엑셀 파일로 일괄 변환해 주는 윈도우 데스크탑 GUI 애플리케이션입니다.

수작업으로 진행하던 데이터 정제 및 그래프 작성 시간을 획기적으로 단축하고, 분석 결과물의 일관성을 높이기 위해 개발되었습니다.

<br/>

## ✨ 주요 기능 (Features)

1. **지능형 데이터 분석 및 필터링**
   - `.xls` 및 `.xlsx` 파일 지원
   - 데이터의 증감량(`Strain_Diff`)을 수학적으로 계산하여 기계의 Loading / Unloading 상태를 정확히 판별
   - Transfer Strain 값이 0 미만인 노이즈 데이터 자동 정제
   
2. **동적 그래프 시각화 (Anti-Collision Labeling)**
   - 추출된 사이클 구간별로 다른 색상의 선 그래프 자동 생성
   - 각 사이클의 최대 응력(Stress) 지점에 도달률(%) 라벨 자동 표기
   - `adjustText` 알고리즘을 적용하여 라벨 간의 겹침 현상 원천 차단

3. **맞춤형 데이터 추출 및 일괄 저장 (Export)**
   - **통합 엑셀:** 원본 데이터, 추출 데이터, 사이클별 요약 데이터를 다중 시트(Multi-sheet)로 저장
   - **분리 엑셀:** 시뮬레이션 및 2차 분석을 위해 `Transfer_Strain`과 `Stress` 데이터만 사이클별로 분할하여 저장
   - **그래프 이미지:** 전체 데이터 및 추출 데이터의 고해상도 PNG 자동 저장
   - **ZIP 일괄 압축:** 생성된 4개의 결과물을 사용자 PC 충돌 없이 내부 임시 폴더에서 생성 후, 하나의 `.zip` 파일로 깔끔하게 압축하여 제공

<br/>

## 🛠 기술 스택 (Tech Stack)

- **Language:** Python 3.x
- **GUI Framework:** Tkinter
- **Data Analysis:** Pandas
- **Data Visualization:** Matplotlib, adjustText
- **Packaging:** PyInstaller, Inno Setup

<br/>

## 🚀 설치 및 실행 방법 (Getting Started)

### 1. 실행 파일(.exe)로 바로 사용하기
코딩 환경 없이 바로 프로그램을 사용하고 싶으시다면, 우측 **[Releases]** 탭에서 최신 버전의 `Setup.exe` 또는 압축 파일을 다운로드하여 실행해 주세요.

### 2. 소스 코드로 직접 실행하기 (개발자용)
```bash
# 저장소 클론
git clone [https://github.com/bbakkomm/biaxial-tension-analyzer.git](https://github.com/bbakkomm/biaxial-tension-analyzer.git)
cd biaxial-tension-analyzer

# 필수 라이브러리 설치
pip install -r requirements.txt

# 프로그램 실행
python app_gui.py