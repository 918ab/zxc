import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import time
import io
import threading
import os 
from datetime import datetime

# Selenium 및 HTML/Excel 관련 라이브러리
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Pandas 설정 (출력 편의를 위함)
pd.set_option('display.width', 1000)
pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', None)

class WebScraperApp:
    
    def _load_settings(self):
        """setting.txt 파일에서 설정값을 로드합니다."""
        settings = {}
        try:
            with open("setting.txt", 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and '=' in line:
                        key, value = line.split('=', 1)
                        settings[key.strip()] = value.strip()
        except FileNotFoundError:
            pass
        except Exception as e:
            pass 
        
        return settings

    def __init__(self, master):
        self.master = master
        master.title("웹 테이블 추출기 (Tkinter)")
        master.geometry("800x750") 
        master.protocol("WM_DELETE_WINDOW", self.on_closing) 

        self.driver = None
        self.all_tables = []
        self.current_table_index = 0 
        self.selection_window = None 
        self.log_text = None 

        # ----------------------------------------------------
        # 1. 설정 파일 로드 및 변수 초기화
        # ----------------------------------------------------
        settings = self._load_settings() 

        # 1-1. 크롬/URL 설정 (경로는 사용자에 맞게 설정되어 있어야 합니다)
        self.user_data_path = tk.StringVar(value=settings.get('user_data_path', r"C:\Users\rmaru\AppData\Local\Google\Chrome\Profile 2"))
        self.profile_dir = tk.StringVar(value=settings.get('profile_dir', "Profile 2"))
        self.target_url = tk.StringVar(value=settings.get('target_url', "https://finance.naver.com/sise/sise_market_sum.nhn"))
        
        # 1-2. 기본 저장 위치 설정 (UI에서 수정 가능)
        self.excel_path = tk.StringVar(value=settings.get('excel_path', r"C:\Users\rmaru\OneDrive\바탕 화면\zxc\dsadsa.xlsx")) 
        self.sheet_name = tk.StringVar(value=settings.get('primary_sheet_name', "테스트")) 
        self.start_row = tk.StringVar(value=settings.get('primary_start_row', "34")) # 저장 시작 행
        
        # 1-3. 보조 저장 위치 설정 (UI에서 수정 가능)
        self.secondary_sheet_name = tk.StringVar(value=settings.get('secondary_sheet_name', "테스트2")) 
        self.secondary_start_row = tk.StringVar(value=settings.get('secondary_start_row', "60")) # 저장 시작 행

        # ----------------------------------------------------
        # 2. UI 레이아웃 구성
        # ----------------------------------------------------
        main_frame = ttk.Frame(master, padding="15")
        main_frame.pack(fill='both', expand=True)

        self._create_setting_section(main_frame, "크롬 프로필 설정", [
            ("User Data Path:", self.user_data_path),
            ("Profile Directory:", self.profile_dir),
            ("Target URL:", self.target_url)
        ])
        
        self._create_excel_section(main_frame)
        
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10, fill='x')
        
        self.main_button = ttk.Button(button_frame, 
                                      text="1. 시작하기", 
                                      command=self.run_open_browser_and_scrape_thread)
        self.main_button.pack(side='left', fill='x', expand=True, padx=5)
        
        self.quit_button = ttk.Button(button_frame, 
                                      text="2. 프로그램 종료", 
                                      command=self.on_closing)
        self.quit_button.pack(side='right', fill='x', expand=True, padx=5)

        self._create_log_section(main_frame) 
        
        self.update_log("프로그램 시작. 설정값을 확인하고 시작하기")
        if settings:
             self.update_log("✅ setting.txt에서 설정값 로드 성공.")
        else:
             self.update_log("❌ setting.txt 파일을 찾을 수 없습니다. 기본값을 사용합니다.")


    def _create_setting_section(self, parent, title, fields):
        labelframe = ttk.LabelFrame(parent, text=title, padding="10")
        labelframe.pack(fill='x', padx=5, pady=5)
        
        for i, (label_text, var) in enumerate(fields):
            ttk.Label(labelframe, text=label_text).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            ttk.Entry(labelframe, textvariable=var, width=60).grid(row=i, column=1, sticky='ew', padx=5, pady=2)

    def _create_excel_section(self, parent):
        labelframe = ttk.LabelFrame(parent, text="엑셀 저장 설정", padding="10")
        labelframe.pack(fill='x', padx=5, pady=5)

        # ----------------------------------------------------
        # 1. Excel 파일 경로
        # ----------------------------------------------------
        ttk.Label(labelframe, text="Excel File Path:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.excel_path, width=40).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Button(labelframe, text="찾아보기", command=self.browse_excel_path).grid(row=0, column=2, sticky='e', padx=5, pady=2)

        # ----------------------------------------------------
        # 2. 기본 저장 위치 (Entry로 수정 가능)
        # ----------------------------------------------------
        ttk.Label(labelframe, text="[기본] Sheet Name:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.sheet_name, width=15).grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(labelframe, text="[기본] Start Row:").grid(row=1, column=1, sticky='e', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.start_row, width=10).grid(row=1, column=2, sticky='e', padx=5, pady=2)
        
        # ----------------------------------------------------
        # 3. 보조 저장 위치 (Entry로 수정 가능)
        # ----------------------------------------------------
        ttk.Label(labelframe, text="[보조] Sheet Name:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.secondary_sheet_name, width=15).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(labelframe, text="[보조] Start Row:").grid(row=2, column=1, sticky='e', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.secondary_start_row, width=10).grid(row=2, column=2, sticky='e', padx=5, pady=2)


    def _create_log_section(self, parent):
        labelframe = ttk.LabelFrame(parent, text="📜 작업 상태 로그", padding="10")
        labelframe.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(labelframe, height=10, state='disabled', wrap='word', bg='#2d2d2d', fg='#d4d4d4', font=('Courier New', 10))
        self.log_text.pack(fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(labelframe, command=self.log_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=scrollbar.set)
        
    def browse_excel_path(self):
        filename = filedialog.askopenfilename(defaultextension=".xlsx",
                                              filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.excel_path.set(filename)

    def update_log(self, message):
        """상태를 로그 형식으로 텍스트 위젯에 추가"""
        
        if self.log_text is None:
            return 
            
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        log_entry = f"{timestamp} {message}\n"
        
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.master.update_idletasks()

    def on_closing(self):
        """프로그램 종료 시 드라이버를 안전하게 닫음"""
        self.update_log("프로그램을 종료합니다.")
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        if self.selection_window and self.selection_window.winfo_exists():
            self.selection_window.destroy()
        self.master.destroy()

    # ----------------------------------------------------
    # 3. Selenium 및 크롤링 로직 
    # ----------------------------------------------------
    
    def run_open_browser_and_scrape_thread(self):
        """GUI가 멈추지 않도록 스레드로 모든 작업을 실행"""
        self.main_button.config(state='disabled', text="⏳ 작업 진행 중...")
        threading.Thread(target=self._integrated_workflow, daemon=True).start()

    def _integrated_workflow(self):
        """브라우저 열기부터 탐색/비교까지의 전체 통합 워크플로우"""
        
        self.update_log("--- 브라우저 열기 및 테이블 탐색 시작 ---")
        
        self.open_browser()
        
        if not self.driver:
            self.main_button.config(state='normal', text="1. 시작하기")
            return

        if not self.all_tables: 
            self.update_log("=========================================================================")
            self.update_log("⚠️ **[중요]** 브라우저 접속 완료. 수동 조작을 완료해 주세요.")
            self.update_log("1. **브라우저에서 직접 '항목 선택' 버튼을 클릭합니다.**")
            self.update_log("2. **필요한 체크박스를 직접 체크하고 '확인'을 누릅니다.**")
            self.update_log("3. **이후, Tkinter 프로그램에서 다시 이 버튼을 클릭하여 탐색을 재개합니다.**")
            self.update_log("=========================================================================")
            self.update_log("5초 대기 후 테이블 탐색을 시도합니다.")
            time.sleep(5) 
            
        self.start_scraping()
        
        self.main_button.config(state='normal', text="1. 시작하기")
    
    def open_browser(self):
        """브라우저를 열고 접속"""
        if self.driver:
            self.update_log("🔄 기존 드라이버 유지, URL 재접속 시도...")
            try:
                self.driver.get(self.target_url.get())
                self.update_log("✅ URL 재접속 성공.")
                return
            except Exception as e:
                self.update_log(f"❌ URL 재접속 오류: {e.__class__.__name__}. 드라이버 재시작 필요.")
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None
        
        self.update_log("⏳ 크롬 브라우저를 열고 지정된 URL로 접속합니다...")
        
        try:
            options = Options()
            options.add_argument(f"user-data-dir={self.user_data_path.get()}") 
            options.add_argument(f"profile-directory={self.profile_dir.get()}") 
            
            # 주의: 사용자의 Chrome 드라이버 경로가 환경변수에 등록되어 있어야 합니다.
            self.driver = webdriver.Chrome(options=options)
            self.driver.get(self.target_url.get())
            
            self.update_log("✅ 브라우저 접속 성공.")
            
        except Exception as e:
            self.update_log(f"❌ 브라우저 오류: {e.__class__.__name__}. 프로필 경로 또는 드라이버 버전을 확인하세요.")
            if self.driver:
                self.driver.quit()
            self.driver = None

    def _restart_scraping(self, current_window):
        """
        테이블 선택 창을 닫고, 테이블 목록을 초기화한 후, 
        스크래핑 로직을 처음부터 다시 시작합니다.
        """
        if current_window:
            current_window.destroy()
        
        self.all_tables = []
        self.current_table_index = 0
        self.update_log("🔄 테이블 목록 초기화 후 스크래핑 로직을 재시작합니다.")
        
        # GUI가 멈추지 않도록 스레드로 메인 워크플로우를 다시 호출합니다.
        self.run_open_browser_and_scrape_thread()


    def start_scraping(self):
        """테이블 탐색 로직"""
        if not self.driver:
             self.update_log("❌ 드라이버가 열려있지 않아 탐색을 시작할 수 없습니다.")
             return
             
        self.update_log("⏳ HTML 소스에서 테이블 탐색을 시작합니다.")
        self.current_table_index = 0
        self.all_tables = []
        
        try:
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            self.update_log("✅ 페이지에서 테이블 요소 감지 완료. 데이터 파싱 중...")
            html_source = self.driver.page_source
            self.all_tables = pd.read_html(io.StringIO(html_source))
            
            if not self.all_tables:
                self.update_log("❌ 테이블을 찾을 수 없습니다. (<table> 태그 없음)")
                messagebox.showerror("오류", "HTML 소스에서 유효한 테이블을 찾지 못했습니다.")
                return

            num_tables = len(self.all_tables)
            self.update_log(f"✅ 총 {num_tables}개의 테이블 발견.")
            
            if num_tables == 1:
                self._open_comparison_window()
            else:
                self.update_log("➡️ 테이블이 여러 개 발견되어, 목록 선택 창을 띄웁니다.")
                self._open_full_selection_window()


        except Exception as e:
            self.update_log(f"❌ 탐색 오류: {e.__class__.__name__}. 상세: {e}")
            messagebox.showerror("오류", f"테이블 탐색 중 오류 발생: {e}")

    # ----------------------------------------------------
    # 4. 엑셀 저장 로직 (ExcelWriter 기반)
    # ----------------------------------------------------

    def _finalize_export(self, df_selected: pd.DataFrame, source_window: tk.Toplevel):
        """
        [ExcelWriter 기반] 최종 요청 로직: 
        1. 기존 엑셀 파일의 모든 시트 내용을 메모리로 읽어옴.
        2. '테스트' 시트의 1~32행 복사 및 데이터 준비.
        3. ExcelWriter를 사용하여 모든 데이터를 지정된 위치에 덮어씀.
        """
        excel_path = self.excel_path.get()
        USER_SHEET_NAME = self.sheet_name.get() 
        FIXED_SHEET_NAME = self.secondary_sheet_name.get()
        
        try:
            USER_START_ROW = int(self.start_row.get())
            FIXED_START_ROW = int(self.secondary_start_row.get())
        except ValueError:
            messagebox.showerror("오류", "시작 행은 유효한 숫자여야 합니다.")
            return

        df_full = df_selected # 웹 데이터 원본
        df_sub = pd.DataFrame() # '테스트' 시트에서 복사될 데이터
        
        # ----------------------------------------------------
        # ⭐️ Inf/NaN 값 처리 (nan_inf_to_erros 방지)
        # ----------------------------------------------------
        self.update_log("⏳ 데이터 클리닝: Inf/NaN 값을 0으로 대체합니다.")
        df_full = df_full.replace([float('inf'), float('-inf')], float('nan'))
        df_full = df_full.fillna(0) 
        self.update_log("✅ 데이터 클리닝 완료.")
        
        # ⚠️ 웹 데이터의 열 정보 
        WEB_COLUMN_COUNT = df_full.shape[1] 
        WEB_COLUMN_NAMES = df_full.columns.tolist() 
        
        # ----------------------------------------------------
        # 1. Task 1: 기존 엑셀 파일의 모든 시트 데이터 읽기
        # ----------------------------------------------------
        existing_sheets = {}
        try:
            # 모든 시트를 한 번에 읽음 (sheet_name=None)
            self.update_log("⏳ Task 1: 기존 엑셀 파일의 모든 시트 데이터 메모리로 로드...")
            # header=None으로 읽어서 모든 데이터를 값으로 처리
            existing_sheets = pd.read_excel(excel_path, sheet_name=None, header=None)
            
        except FileNotFoundError:
            self.update_log("⚠️ 기존 파일이 없어 새로 생성합니다.")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 로드 중 오류 발생: {e.__class__.__name__}\n상세: {e}")
            self.update_log(f"❌ 기존 파일 로드 실패: {e.__class__.__name__}. 상세: {e}")
            return
            
        # ----------------------------------------------------
        # 2. Task 4: '테스트' 시트 (1~32행) 복사 (⭐️ 모든 열 복사)
        # ----------------------------------------------------
        
        if USER_SHEET_NAME in existing_sheets:
            self.update_log(f"⏳ Task 4: '{USER_SHEET_NAME}' 시트의 1~32행 복사 준비 중...")
            
            df_test_full = existing_sheets[USER_SHEET_NAME]
            
            # 1행(인덱스 0)부터 32행(인덱스 31)까지 복사
            try:
                # ⭐️ 모든 열을 가져옵니다. (A부터 끝까지)
                df_from_excel = df_test_full.iloc[0:32, :] 
                
                df_sub = df_from_excel.copy()
                
                # 복사된 데이터의 열 이름을 웹 데이터의 열 개수에 맞게 조정
                if df_sub.shape[1] > WEB_COLUMN_COUNT:
                    # 웹 데이터 열 이름 개수를 초과하는 나머지 열 이름 생성
                    remaining_cols = [f'Unnamed_{i}' for i in range(WEB_COLUMN_COUNT, df_sub.shape[1])]
                    new_cols = WEB_COLUMN_NAMES + remaining_cols
                    df_sub.columns = new_cols[:df_sub.shape[1]]
                else:
                    # 복사된 데이터의 열 개수가 더 적거나 같으면 웹 데이터 열 이름만 적용
                    df_sub.columns = WEB_COLUMN_NAMES[:df_sub.shape[1]]
                
                self.update_log(f"   ✅ '{USER_SHEET_NAME}' 시트의 1~32행 복사 성공. (크기: {df_sub.shape})")
            except IndexError:
                self.update_log(f"⚠️ '{USER_SHEET_NAME}' 시트의 데이터가 충분하지 않아 복사를 건너킵니다.")
        else:
            self.update_log(f"⚠️ '{USER_SHEET_NAME}' 시트가 없어 복사할 데이터(1~32행)는 생성되지 않습니다.")
            
        # ----------------------------------------------------
        # 3. Task 1-2 & 5 & 6: ExcelWriter를 사용한 데이터 통합 및 저장
        # ----------------------------------------------------
        
        self.update_log("⏳ Task 1-2 & 5: ExcelWriter를 사용하여 데이터 통합 및 저장 시작...")
        
        try:
            # writer 생성 (파일을 덮어쓰기 모드 'w'로 엽니다)
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                
                # 엑셀 워크북과 포맷 준비
                workbook = writer.book
                
                # 테두리가 없는 기본 포맷 정의
                no_border_format = workbook.add_format({
                    'border': 0, 'top': 0, 'bottom': 0, 'left': 0, 'right': 0,
                    'align': 'center', 'valign': 'vcenter' 
                })
                
                # 3-1. 기존 시트 데이터 먼저 쓰기 (복사된 시트를 위해)
                for sheet_name, df_data in existing_sheets.items():
                    # '테스트'와 '테스트2'는 아래에서 덮어쓸 예정이므로 제외
                    if sheet_name not in [USER_SHEET_NAME, FIXED_SHEET_NAME]:
                        # 기존 데이터는 A1(row=0, col=0)부터 쓰기
                        df_data.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, header=False, index=False)
                        self.update_log(f"   ... 기존 시트 '{sheet_name}' 저장 완료.")


                # 3-2. '테스트' 시트 데이터 저장 (기존 데이터 + 웹 데이터)
                
                worksheet = workbook.add_worksheet(USER_SHEET_NAME)
                writer.sheets[USER_SHEET_NAME] = worksheet 
                
                # a. 기존 1~33행 데이터 쓰기 (Pandas index 0~32)
                if USER_SHEET_NAME in existing_sheets:
                    df_existing_top = existing_sheets[USER_SHEET_NAME].iloc[0:USER_START_ROW-1]
                    df_existing_top.to_excel(writer, 
                                            sheet_name=USER_SHEET_NAME, 
                                            startrow=0, startcol=0, 
                                            header=False, index=False) 
                    self.update_log(f"   ... '{USER_SHEET_NAME}' 시트 기존 데이터 (1행부터 {USER_START_ROW-1}행) 저장 완료.")
                
                # b. 웹 데이터 쓰기 (34행부터) - 'write_row' 메서드 강제 사용
                start_row_excel = USER_START_ROW - 1 # 엑셀 34행 = Pandas index 33
                
                # 헤더를 no_border_format으로 강제 쓰기
                header_list = df_full.columns.tolist()
                worksheet.write_row(start_row_excel, 0, header_list, no_border_format)
                self.update_log(f"   ✅ 웹 데이터 헤더 -> '{USER_SHEET_NAME}' 시트 ({USER_START_ROW}행) 테두리 없이 저장 완료.")
                
                # 데이터 본체를 행 단위로 순회하며 no_border_format으로 쓰기
                for row_index, row_data in df_full.iterrows():
                    excel_row = start_row_excel + 1 + row_index
                    worksheet.write_row(excel_row, 0, row_data.tolist(), no_border_format)
                
                self.update_log(f"   ✅ 웹 데이터 본문 -> '{USER_SHEET_NAME}' 시트 ({start_row_excel+2}~행) 테두리 없이 저장 완료.")

                
                # 3-3. '테스트2' 시트 데이터 저장
                # Task 5: 복사 데이터 -> '테스트2' 시트 (60행)
                if not df_sub.empty:
                    df_sub.to_excel(writer, 
                                    sheet_name=FIXED_SHEET_NAME, 
                                    startrow=FIXED_START_ROW - 1, startcol=0, 
                                    header=False, index=False) 
                    self.update_log(f"   ✅ 복사 데이터 -> '{FIXED_SHEET_NAME}' 시트 ({FIXED_START_ROW}행) 저장 완료 (헤더 제외).")
                else:
                    self.update_log(f"⚠️ 복사할 데이터가 없어 '{FIXED_SHEET_NAME}' 시트 저장은 건너킵니다.")


            self.update_log(f"🎉 **최종 저장 완료:** '{excel_path}' 파일에 모든 작업이 안전하게 반영되었습니다.")
            source_window.destroy()

        except PermissionError:
            self.update_log(f"❌ 파일 잠금 오류! '{excel_path}' 파일이 열려있습니다. 임시 저장 로직으로 이동합니다.")
            self._handle_temp_save(df_full, source_window)
            return
        except Exception as e:
            messagebox.showerror("오류", f"ExcelWriter 저장 중 치명적인 오류 발생: {e.__class__.__name__}\n상세: {e}")
            self.update_log(f"❌ ExcelWriter 저장 실패: {e.__class__.__name__}. 상세: {e}")
            return
            
            
    def _handle_temp_save(self, df_full, source_window):
        """파일 잠금 오류 발생 시 임시 파일에 저장하는 로직"""
        excel_path = self.excel_path.get()
        USER_SHEET_NAME = self.sheet_name.get()
        FIXED_SHEET_NAME = self.secondary_sheet_name.get()
        
        try:
            USER_START_ROW = int(self.start_row.get())
            FIXED_START_ROW = int(self.secondary_start_row.get())
        except ValueError:
            messagebox.showerror("오류", "시작 행은 유효한 숫자여야 합니다.")
            return

        # 임시 파일 경로 생성
        base, ext = os.path.splitext(excel_path)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_path = f"{base}_TEMP_{timestamp}{ext}"
        
        # ----------------------------------------------------
        # 임시 저장 시에도 기존 파일 로드 시도
        # ----------------------------------------------------
        existing_sheets = {}
        df_sub_temp = pd.DataFrame()
        
        try:
            # 기존 파일 로드 시도 (header=None)
            existing_sheets = pd.read_excel(excel_path, sheet_name=None, header=None)
            
            # 테스트 시트의 1~32행 복사
            df_test_full = existing_sheets.get(USER_SHEET_NAME, pd.DataFrame())
            if not df_test_full.empty and len(df_test_full) >= 32:
                # 1행(인덱스 0)부터 32행(인덱스 31)까지 복사 (모든 열 포함)
                df_from_excel = df_test_full.iloc[0:32, :]
                
                df_sub_temp = df_from_excel.copy()
                
                # 웹 데이터 열 정보 다시 가져오기
                WEB_COLUMN_COUNT = df_full.shape[1] 
                WEB_COLUMN_NAMES = df_full.columns.tolist() 
                
                # 복사된 데이터의 열 이름을 웹 데이터의 열 개수에 맞게 조정
                if df_sub_temp.shape[1] > WEB_COLUMN_COUNT:
                    remaining_cols = [f'Unnamed_{i}' for i in range(WEB_COLUMN_COUNT, df_sub_temp.shape[1])]
                    new_cols = WEB_COLUMN_NAMES + remaining_cols
                    df_sub_temp.columns = new_cols[:df_sub_temp.shape[1]]
                else:
                    df_sub_temp.columns = WEB_COLUMN_NAMES[:df_sub_temp.shape[1]]
                
            else:
                 # 파일 로드 실패 시, 복사 데이터는 웹 데이터의 상위 32행이라고 가정
                 df_sub_temp = df_full.head(32).copy()
        except:
             # 파일 로드 실패 시, 복사 데이터는 웹 데이터의 상위 32행이라고 가정
             df_sub_temp = df_full.head(32).copy()
             
        self.update_log(f"   - 임시 저장 파일: {temp_path}")
        
        try:
            with pd.ExcelWriter(temp_path, engine='xlsxwriter') as writer:
                
                # 엑셀 워크북과 포맷 준비 (임시 저장에서도 동일하게 적용)
                workbook = writer.book
                no_border_format = workbook.add_format({
                    'border': 0, 'top': 0, 'bottom': 0, 'left': 0, 'right': 0,
                    'align': 'center', 'valign': 'vcenter' 
                })
                
                # 1. 기존 시트 데이터 쓰기 (테스트, 테스트2 제외)
                for sheet_name, df_data in existing_sheets.items():
                    if sheet_name not in [USER_SHEET_NAME, FIXED_SHEET_NAME]:
                        df_data.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, header=False, index=False)
                        
                # 2. '테스트' 시트 저장 (write_row 방식)
                worksheet = workbook.add_worksheet(USER_SHEET_NAME)
                writer.sheets[USER_SHEET_NAME] = worksheet 
                
                # 기존 1행부터 33행 데이터 쓰기
                if USER_SHEET_NAME in existing_sheets:
                    df_existing_top = existing_sheets[USER_SHEET_NAME].iloc[0:USER_START_ROW-1]
                    df_existing_top.to_excel(writer, 
                                            sheet_name=USER_SHEET_NAME, 
                                            startrow=0, startcol=0, 
                                            header=False, index=False)
                                            
                start_row_excel = USER_START_ROW - 1 # 엑셀 34행
                
                # 헤더를 no_border_format으로 강제 쓰기
                header_list = df_full.columns.tolist()
                worksheet.write_row(start_row_excel, 0, header_list, no_border_format)
                
                # 데이터 본체를 행 단위로 순회하며 no_border_format으로 쓰기
                for row_index, row_data in df_full.iterrows():
                    excel_row = start_row_excel + 1 + row_index
                    worksheet.write_row(excel_row, 0, row_data.tolist(), no_border_format)
                                                
                # 3. '테스트2' 시트 저장
                if not df_sub_temp.empty:
                    df_sub_temp.to_excel(writer, 
                                         sheet_name=FIXED_SHEET_NAME, 
                                         startrow=FIXED_START_ROW - 1, startcol=0, 
                                         header=False, index=False)
            
            error_message = f"✅ **임시 저장 완료!**\n\n**[사유]** 원래 파일이 열려있어 데이터를 임시 파일에 저장했습니다.\n\n1. 원래 파일 ('{excel_path}')을 **닫아주세요.**\n2. **'{temp_path}'** 파일을 열어 내용을 복사해서 원래 파일에 덮어씌워 주세요. (두 시트 모두 임시 파일에 있습니다.)"
            self.update_log(error_message)
            messagebox.showinfo("🚨 임시 저장 완료", error_message)
            source_window.destroy()

        except Exception as temp_e:
            self.update_log(f"❌ 임시 저장 중에도 오류 발생: {temp_e.__class__.__name__}")
            messagebox.showerror("오류", f"임시 파일 저장 중 치명적인 오류 발생: {temp_e}")


# -------------------------------------------------------------------------------------
# (테이블 선택 및 비교를 위한 GUI 팝업 함수 - 변경 없음)
# -------------------------------------------------------------------------------------
def _open_comparison_window_impl(app_instance):
    
    if app_instance.selection_window and app_instance.selection_window.winfo_exists():
        app_instance.selection_window.destroy()

    if app_instance.current_table_index >= len(app_instance.all_tables):
        app_instance.update_log("⚠️ 현재 테이블 인덱스에서 더 이상 테이블을 찾을 수 없습니다. (목록 끝)")
        messagebox.showinfo("정보", "현재 테이블 탐색 목록의 끝에 도달했습니다. 재탐색을 다시 시작하거나 URL을 확인하세요.")
        return
        
    df = app_instance.all_tables[app_instance.current_table_index]
    df_cleaned = df.dropna(how='all')
    
    selection_window = tk.Toplevel(app_instance.master)
    selection_window.title(f"테이블 비교: #{app_instance.current_table_index + 1} / {len(app_instance.all_tables)}개")
    selection_window.geometry("800x500")
    
    ttk.Label(selection_window, 
              text=f"현재 테이블 #{app_instance.current_table_index + 1}의 미리보기입니다. 이 테이블이 맞습니까?", 
              font=('Arial', 12, 'bold')).pack(pady=10)
    
    ttk.Label(selection_window, 
              text=f"크기: {df_cleaned.shape[0]}행 x {df_cleaned.shape[1]}열", 
              font=('Arial', 10)).pack(pady=5)
    
    preview_frame = ttk.Frame(selection_window, borderwidth=2, relief="groove")
    preview_frame.pack(padx=10, pady=5, fill='both', expand=True)
    
    preview_text = tk.Text(preview_frame, height=15, width=90, font=('Courier New', 9), wrap='none')
    preview_text.pack(side='left', fill='both', expand=True)
    
    preview_content = df_cleaned.head(10).to_string(header=True, index=False)
    preview_text.insert(tk.END, preview_content)
    preview_text.config(state='disabled')
    
    v_scroll = ttk.Scrollbar(preview_frame, command=preview_text.yview)
    v_scroll.pack(side='right', fill='y')
    preview_text.config(yscrollcommand=v_scroll.set)

    button_frame = ttk.Frame(selection_window)
    button_frame.pack(pady=15)
    
    ttk.Button(button_frame, text="✅ 이 테이블이 맞습니다 (엑셀 저장)", 
               command=lambda: app_instance._finalize_export(df_cleaned, selection_window)).pack(side='left', padx=10)
    
    ttk.Button(button_frame, text="⏭️ 이 테이블이 아님 (다음 테이블 보기)", 
               command=lambda: _move_to_next_table_impl(app_instance, selection_window)).pack(side='left', padx=10)
    
    app_instance.selection_window = selection_window
    selection_window.transient(app_instance.master)
    app_instance.master.wait_window(selection_window)

def _move_to_next_table_impl(app_instance, current_window):
    
    if current_window:
        current_window.destroy()
    app_instance.current_table_index += 1
    app_instance.update_log(f"↪️ 테이블 #{app_instance.current_table_index}를 건너뛰고 다음 테이블 탐색을 요청합니다.")
    app_instance.master.after(10, lambda: _open_comparison_window_impl(app_instance)) 

def _open_full_selection_window_impl(app_instance):
    
    if app_instance.selection_window and app_instance.selection_window.winfo_exists():
        app_instance.selection_window.destroy()
        
    selection_window = tk.Toplevel(app_instance.master)
    selection_window.title(f"전체 테이블 목록에서 선택 (총 {len(app_instance.all_tables)}개)")
    selection_window.geometry("900x700") 
    
    ttk.Label(selection_window, text="👀 테이블이 여러 개 발견되었습니다. 목록에서 하나를 선택하고 저장하거나, 재탐색하세요.", 
              font=('Arial', 12, 'bold'), foreground='blue').pack(pady=10)
    
    canvas = tk.Canvas(selection_window)
    scrollbar = ttk.Scrollbar(selection_window, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")
    
    app_instance.df_buttons = [] 
    
    for i, df in enumerate(app_instance.all_tables):
        df_cleaned = df.dropna(how='all')
        
        table_frame = ttk.LabelFrame(scrollable_frame, text=f"테이블 #{i+1} (크기: {df_cleaned.shape[0]}행 x {df_cleaned.shape[1]}열)", padding="10")
        table_frame.pack(fill='x', padx=5, pady=5)
        
        preview_text = tk.Text(table_frame, height=5, width=100, font=('Courier New', 9), wrap='none', state='normal')
        preview_content = df_cleaned.head(5).to_string(header=True, index=False)
        preview_text.insert(tk.END, preview_content)
        preview_text.config(state='disabled')
        preview_text.pack(fill='x', pady=5)

        select_button = ttk.Button(table_frame, text="✅ 이 테이블 선택 및 저장", 
                                   command=lambda d=df_cleaned: app_instance._finalize_export(d, selection_window))
        select_button.pack(side='right', pady=5)
        app_instance.df_buttons.append(df_cleaned)

    bottom_button_frame = ttk.Frame(selection_window)
    bottom_button_frame.pack(pady=10)
    
    ttk.Button(bottom_button_frame, 
               text="🔄 재탐색 (테이블 다시 인식)", 
               command=lambda: app_instance._restart_scraping(selection_window)).pack(padx=10)
    
    app_instance.selection_window = selection_window
    selection_window.transient(app_instance.master)
    app_instance.master.wait_window(selection_window)

# 메인 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = WebScraperApp(root)
    # 클래스 메서드에 외부 함수 연결
    WebScraperApp._open_comparison_window = _open_comparison_window_impl
    WebScraperApp._move_to_next_table = _move_to_next_table_impl
    WebScraperApp._open_full_selection_window = _open_full_selection_window_impl

    root.mainloop()