import matplotlib
matplotlib.use('TkAgg') 
import matplotlib.pyplot as plt
from adjustText import adjust_text 

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import pandas as pd
import os
import sys
import datetime
import zipfile   # ZIP 압축을 위한 모듈 추가
import tempfile  # 임시 폴더 생성을 위한 모듈 추가

class BiaxialAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("고무벨트 인장시험 데이터 분석기 (최종 납품용)")
        self.root.geometry("750x650")
        
        self.filepath = ""
        self.df_raw = None
        self.df_filtered = None
        self.max_cycle = 0
        
        self.setup_ui()
        self.log("프로그램이 시작되었습니다. 엑셀 파일을 불러와주세요.")

    def setup_ui(self):
        style = ttk.Style()
        style.configure('TButton', font=('맑은 고딕', 10))
        style.configure('TLabel', font=('맑은 고딕', 10))
        
        # 1. 파일 불러오기
        frame_top = ttk.LabelFrame(self.root, text="1. 데이터 불러오기", padding=10)
        frame_top.pack(fill='x', padx=10, pady=5)
        
        self.lbl_file = ttk.Label(frame_top, text="선택된 파일: 없음", foreground="blue")
        self.lbl_file.pack(side='left', fill='x', expand=True)
        btn_load = ttk.Button(frame_top, text="파일 찾기", command=self.load_file)
        btn_load.pack(side='right')

        # 2. 분석 및 추출
        frame_mid = ttk.LabelFrame(self.root, text="2. 사이클(Cycle) 분석 및 추출", padding=10)
        frame_mid.pack(fill='x', padx=10, pady=5)
        
        self.lbl_cycle_info = ttk.Label(frame_mid, text="총 반복 횟수: (파일을 먼저 불러오세요)")
        self.lbl_cycle_info.grid(row=0, column=0, columnspan=3, sticky='w', pady=5)
        
        ttk.Label(frame_mid, text="추출할 사이클 번호 (쉼표 구분):").grid(row=1, column=0, sticky='w')
        self.entry_cycles = ttk.Entry(frame_mid, width=25)
        self.entry_cycles.grid(row=1, column=1, padx=5)
        self.entry_cycles.insert(0, "5, 10, 15")
        
        btn_apply = ttk.Button(frame_mid, text="데이터 추출 적용", command=self.apply_filter)
        btn_apply.grid(row=1, column=2, padx=5)

        # 3. 결과 확인 및 저장
        frame_actions = ttk.LabelFrame(self.root, text="3. 결과 확인 및 개별 저장", padding=10)
        frame_actions.pack(fill='x', padx=10, pady=5)
        
        self.action_buttons = []

        frame_charts = ttk.Frame(frame_actions)
        frame_charts.pack(fill='x', pady=5)
        btn_chart1 = ttk.Button(frame_charts, text="전체 그래프 보기", command=self.show_raw_chart)
        btn_chart2 = ttk.Button(frame_charts, text="추출 그래프 보기", command=self.show_filtered_chart)
        btn_chart1.pack(side='left', padx=5, expand=True, fill='x')
        btn_chart2.pack(side='left', padx=5, expand=True, fill='x')
        self.action_buttons.extend([btn_chart1, btn_chart2])

        frame_saves = ttk.Frame(frame_actions)
        frame_saves.pack(fill='x', pady=5)
        
        btn_save_excel = ttk.Button(frame_saves, text="엑셀 결과 (시트 통합)", command=self.save_excel_integrated)
        btn_save_cycle = ttk.Button(frame_saves, text="사이클별 분리 엑셀", command=self.save_cycle_separated_excel)
        btn_save_c1 = ttk.Button(frame_saves, text="전체 그래프 저장", command=lambda: self.save_chart('raw'))
        btn_save_c2 = ttk.Button(frame_saves, text="추출 그래프 저장", command=lambda: self.save_chart('filtered'))
        
        for btn in [btn_save_excel, btn_save_cycle, btn_save_c1, btn_save_c2]:
            btn.pack(side='left', padx=2, expand=True, fill='x')
            self.action_buttons.append(btn)

        # 4. [변경] 일괄 압축 저장 버튼 (시각적으로 띄우기)
        frame_zip = ttk.Frame(self.root)
        frame_zip.pack(fill='x', padx=10, pady=10)
        btn_save_all = ttk.Button(frame_zip, text="📦 결과물 일괄 저장 (ZIP 압축 다운로드)", command=self.save_all_to_zip)
        btn_save_all.pack(fill='x', ipady=5) # ipady로 버튼 높이를 약간 키움
        self.action_buttons.append(btn_save_all)

        self.toggle_action_buttons("disable")

        # 5. 로그
        frame_log = ttk.LabelFrame(self.root, text="진행 상황 및 오류 메시지", padding=10)
        frame_log.pack(fill='both', expand=True, padx=10, pady=5)
        self.txt_log = scrolledtext.ScrolledText(frame_log, height=10, font=('맑은 고딕', 9))
        self.txt_log.pack(fill='both', expand=True)

    def log(self, message):
        self.txt_log.insert(tk.END, message + "\n")
        self.txt_log.see(tk.END)
        self.root.update()

    def toggle_action_buttons(self, state):
        for btn in self.action_buttons:
            btn.config(state=state)

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath: return

        self.filepath = filepath
        self.lbl_file.config(text=f"선택된 파일: {os.path.basename(filepath)}")
        self.log(f"\n[파일 읽기] '{os.path.basename(filepath)}' 분석을 시작합니다...")

        try:
            df = pd.read_excel(filepath)
            if 'Strain' not in df.columns or 'Stress' not in df.columns:
                raise ValueError("엑셀 파일에 'Strain' 또는 'Stress' 열이 존재하지 않습니다.")
            
            df['Transfer Strain'] = df['Strain'] / 100
            df['Strain Diff'] = df['Transfer Strain'].diff()
            df['State'] = df['Strain Diff'].apply(lambda x: 'Loading' if x > 0 else 'Unloading')
            df.loc[0, 'State'] = 'Loading'
            
            prev_state = df['State'].shift(1)
            cycle_start = ((prev_state == 'Unloading') & (df['State'] == 'Loading')).astype(int)
            df['Cycle'] = cycle_start.cumsum() + 1
            
            self.df_raw = df
            self.max_cycle = df['Cycle'].max()
            self.lbl_cycle_info.config(text=f"총 반복 횟수: {self.max_cycle} 사이클 감지됨")
            self.log(f"[완료] 총 {self.max_cycle}개의 사이클이 정상적으로 분석되었습니다.")
            
            self.apply_filter()
            
        except Exception as e:
            self.log(f"[오류 발생] 파일을 처리하는 중 문제가 발생했습니다.\n상세내용: {e}")
            messagebox.showerror("파일 읽기 오류", f"데이터 분석 중 오류가 발생했습니다.\n{e}")

    def apply_filter(self):
        if self.df_raw is None:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 불러와주세요.")
            return

        input_str = self.entry_cycles.get()
        try:
            target_cycles = [int(c.strip()) for c in input_str.split(',')]
            invalid_cycles = [c for c in target_cycles if c > self.max_cycle or c < 1]
            if invalid_cycles:
                raise ValueError(f"입력하신 사이클({invalid_cycles})이 전체 반복 횟수({self.max_cycle}) 범위를 벗어납니다.")
        except ValueError as e:
            self.log(f"[입력 오류] {e}")
            messagebox.showerror("입력 오류", "정확한 사이클 번호를 숫자로 입력해주세요.")
            return

        temp_df = self.df_raw[(self.df_raw['State'] == 'Loading') & (self.df_raw['Cycle'].isin(target_cycles))]
        
        original_len = len(temp_df)
        temp_df = temp_df[temp_df['Transfer Strain'] >= 0]
        removed_count = original_len - len(temp_df)
        
        if removed_count > 0:
            self.log(f"[데이터 정제] Transfer Strain이 0 미만인 데이터 {removed_count}개를 자동으로 제거했습니다.")

        self.df_filtered = temp_df
        self.log(f"[추출 완료] {target_cycles}회차 데이터 추출 완료. 결과를 확인하거나 저장할 수 있습니다.")
        self.toggle_action_buttons("normal")

    def setup_plt_font(self):
        plt.rcParams['font.family'] = 'Malgun Gothic'
        plt.rcParams['axes.unicode_minus'] = False

    def get_raw_chart(self):
        self.setup_plt_font()
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.plot(self.df_raw['Transfer Strain'], self.df_raw['Stress'], color='royalblue', linewidth=1)
        ax.set_title('동일고무벨트_Biaxial Tension (전체 데이터)')
        ax.set_xlabel('Nominal Strain')
        ax.set_ylabel('Nominal Stress (MPa)')
        ax.grid(True, linestyle='--', alpha=0.7)
        return fig

    def get_filtered_chart(self):
        self.setup_plt_font()
        fig, ax = plt.subplots(figsize=(8, 5))
        
        target_cycles = sorted(self.df_filtered['Cycle'].unique())
        texts = [] 
        
        for i, cycle in enumerate(target_cycles):
            cycle_data = self.df_filtered[self.df_filtered['Cycle'] == cycle]
            lines = ax.plot(cycle_data['Transfer Strain'], cycle_data['Stress'], 
                             linewidth=1.5, label=f'{cycle}회차')
            
            current_color = lines[0].get_color()
            
            if not cycle_data.empty:
                last_point = cycle_data.iloc[-1]
                x_end = last_point['Transfer Strain']
                y_end = last_point['Stress']
                
                strain_percent = round(x_end * 100)
                text_label = f"{cycle} Cycle ({strain_percent}%)" 
                
                txt = ax.text(x_end, y_end, text_label, 
                              color=current_color,
                              fontsize=9, fontweight='bold')
                texts.append(txt)
        
        if texts:
            self.log("-> 그래프 라벨 위치 최적화 중...")
            self.root.update()
            adjust_text(texts, 
                        add_objects=ax.lines, 
                        autoalign='xy',
                        arrowprops=dict(arrowstyle='-', color='gray', lw=0.5))
            self.log("-> 최적화 완료.")

        ax.set_title(f'동일고무벨트_Biaxial (추출된 {len(target_cycles)}개 구간)')
        ax.set_xlabel('Engineering Strain')
        ax.set_ylabel('Engineering Stress (MPa)')
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.margins(x=0.2, y=0.1) 
        plt.tight_layout()
        
        return fig

    def show_raw_chart(self):
        fig = self.get_raw_chart()
        plt.show()

    def show_filtered_chart(self):
        fig = self.get_filtered_chart()
        plt.show()

    # --- 개별 저장 함수들 (우회 방어 로직 포함) ---
    def save_excel_integrated(self, auto_dir=None):
        default_name = "1_통합_데이터_분석결과.xlsx"
        save_path = os.path.join(auto_dir, default_name) if auto_dir else filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel files", "*.xlsx")])
        
        if not save_path: return

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                self.df_raw.to_excel(writer, sheet_name='1_Raw_Data', index=False)
                self.df_filtered.to_excel(writer, sheet_name='2_Filtered_Data', index=False)
                summary_df = self.df_filtered.groupby('Cycle').tail(1)
                summary_df.to_excel(writer, sheet_name='3_Cycle_Summary', index=False)
            if not auto_dir: self.log(f"[저장 성공] 통합 엑셀 파일 -> {os.path.basename(save_path)}")
        except Exception as e:
            self.log("-> ⚠️ 기존 파일이 열려있거나 잠겨있습니다. 우회 저장을 시도합니다...")
            timestamp = datetime.datetime.now().strftime("%H%M%S")
            base, ext = os.path.splitext(save_path)
            bypass_path = f"{base}_우회저장_{timestamp}{ext}"
            try:
                with pd.ExcelWriter(bypass_path, engine='openpyxl') as writer:
                    self.df_raw.to_excel(writer, sheet_name='1_Raw_Data', index=False)
                    self.df_filtered.to_excel(writer, sheet_name='2_Filtered_Data', index=False)
                    summary_df = self.df_filtered.groupby('Cycle').tail(1)
                    summary_df.to_excel(writer, sheet_name='3_Cycle_Summary', index=False)
                if not auto_dir: self.log(f"✅ [우회 저장 성공] 새 이름으로 안전하게 저장되었습니다 -> {os.path.basename(bypass_path)}")
            except Exception as e2:
                self.log(f"❌ [저장 실패] 우회 저장 실패: {e2}")

    def save_cycle_separated_excel(self, auto_dir=None):
        if self.df_filtered is None: return

        default_name = "2_사이클별_분리데이터.xlsx"
        save_path = os.path.join(auto_dir, default_name) if auto_dir else filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel files", "*.xlsx")])
        
        if not save_path: return

        def write_excel(path):
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                target_cycles = sorted(self.df_filtered['Cycle'].unique())
                for cycle in target_cycles:
                    cycle_df = self.df_filtered[self.df_filtered['Cycle'] == cycle]
                    export_df = cycle_df[['Transfer Strain', 'Stress']]
                    export_df.to_excel(writer, sheet_name=f'{cycle}_Cycle', index=False)

        try:
            write_excel(save_path)
            if not auto_dir: self.log(f"[저장 성공] 사이클별 분리 엑셀 파일 -> {os.path.basename(save_path)}")
        except Exception as e:
            self.log("-> ⚠️ 기존 파일이 열려있거나 잠겨있습니다. 우회 저장을 시도합니다...")
            timestamp = datetime.datetime.now().strftime("%H%M%S")
            base, ext = os.path.splitext(save_path)
            bypass_path = f"{base}_우회저장_{timestamp}{ext}"
            try:
                write_excel(bypass_path)
                if not auto_dir: self.log(f"✅ [우회 저장 성공] 새 이름으로 안전하게 저장되었습니다 -> {os.path.basename(bypass_path)}")
            except Exception as e2:
                self.log(f"❌ [저장 실패] 사이클별 분리 엑셀 파일: {e2}")

    def save_chart(self, chart_type, auto_dir=None):
        if chart_type == 'raw':
            fig = self.get_raw_chart()
            default_name = "3_전체데이터_그래프.png"
        else:
            fig = self.get_filtered_chart()
            default_name = "4_추출데이터_그래프.png"

        save_path = os.path.join(auto_dir, default_name) if auto_dir else filedialog.asksaveasfilename(
            defaultextension=".png", initialfile=default_name, filetypes=[("PNG Image", "*.png")])

        if save_path:
            try:
                fig.savefig(save_path, dpi=300, bbox_inches='tight')
                plt.close(fig)
                if not auto_dir: self.log(f"[저장 성공] 이미지 저장 -> {os.path.basename(save_path)}")
            except Exception as e:
                self.log("-> ⚠️ 기존 이미지 덮어쓰기 실패. 우회 저장을 시도합니다...")
                timestamp = datetime.datetime.now().strftime("%H%M%S")
                base, ext = os.path.splitext(save_path)
                bypass_path = f"{base}_우회저장_{timestamp}{ext}"
                try:
                    fig.savefig(bypass_path, dpi=300, bbox_inches='tight')
                    plt.close(fig)
                    if not auto_dir: self.log(f"✅ [우회 저장 성공] 새 이름으로 이미지 저장 -> {os.path.basename(bypass_path)}")
                except Exception as e2:
                    self.log(f"❌ [저장 실패] 이미지: {e2}")

    # --- [핵심 기능] 일괄 ZIP 압축 저장 ---
    def save_all_to_zip(self):
        """결과물 4개를 보이지 않는 임시 폴더에 만들고, 이를 하나의 ZIP 파일로 압축하여 저장합니다."""
        
        # 1. 사용자에게 저장할 ZIP 파일 이름과 위치를 묻습니다.
        zip_path = filedialog.asksaveasfilename(
            title="모든 결과물을 ZIP 압축 파일로 저장",
            defaultextension=".zip",
            initialfile="고무벨트_분석결과_전체.zip",
            filetypes=[("ZIP Archive", "*.zip")]
        )
        if not zip_path: return

        self.log("\n[일괄 저장 시작] 결과물 4개를 생성하고 ZIP으로 압축 중입니다...")
        
        try:
            # 2. 파이썬이 안전한 임시 폴더를 생성합니다. (작업 후 찌꺼기 없이 자동 삭제됨)
            with tempfile.TemporaryDirectory() as temp_dir:
                
                # 3. 임시 폴더(temp_dir)에 4개의 파일을 조용히 저장합니다.
                self.save_excel_integrated(auto_dir=temp_dir)
                self.save_cycle_separated_excel(auto_dir=temp_dir)
                self.save_chart('raw', auto_dir=temp_dir)
                self.save_chart('filtered', auto_dir=temp_dir)
                
                # 4. 임시 폴더에 생성된 파일들을 하나로 묶어 ZIP 파일을 만듭니다.
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for folder_name, subfolders, filenames in os.walk(temp_dir):
                        for filename in filenames:
                            # 개별 파일의 절대 경로
                            file_path = os.path.join(folder_name, filename)
                            # ZIP 파일 내부에 들어갈 이름
                            zipf.write(file_path, arcname=filename)
                            
            self.log(f"🎉 모든 결과물이 1개의 파일로 압축되어 저장되었습니다! -> {os.path.basename(zip_path)}")
            messagebox.showinfo("압축 저장 완료", "통합 엑셀, 분리 엑셀, 그래프 2종이 성공적으로 ZIP 파일로 압축되었습니다.")
            
        except Exception as e:
            self.log(f"❌ [압축 저장 실패] 오류가 발생했습니다: {e}")
            messagebox.showerror("저장 오류", f"ZIP 파일 저장 중 오류가 발생했습니다.\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = BiaxialAnalyzerApp(root)
    root.mainloop()