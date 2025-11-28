import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import time
import io
import threading
import os 
from datetime import datetime

# Selenium 및 라이브러리
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select 
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Pandas 설정
pd.set_option('display.width', 1000)
pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', None)

class WebScraperApp:
    
    def _load_settings(self):
        settings = {}
        try:
            with open("setting.txt", 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and '=' in line:
                        key, value = line.split('=', 1)
                        settings[key.strip()] = value.strip()
        except: pass 
        return settings

    def __init__(self, master):
        self.master = master
        master.title("웹 테이블 추출기 v5.0 (물류센터 맞춤형)")
        master.geometry("950x850") 
        master.protocol("WM_DELETE_WINDOW", self.on_closing) 

        self.style = ttk.Style()
        self.style.theme_use('clam') 
        self.style.configure('Green.TButton', font=('Malgun Gothic', 10, 'bold'), background='#28a745', foreground='white', borderwidth=1)
        self.style.map('Green.TButton', background=[('active', '#218838')])
        self.style.configure('Blue.TButton', font=('Malgun Gothic', 10, 'bold'), background='#007bff', foreground='white', borderwidth=1)
        self.style.map('Blue.TButton', background=[('active', '#0069d9')])
        self.style.configure('Red.TButton', font=('Malgun Gothic', 9), background='#dc3545', foreground='white')
        self.style.map('Red.TButton', background=[('active', '#c82333')])

        self.driver = None
        self.all_tables = []
        self.current_table_index = 0 
        self.selection_window = None 
        self.log_text = None 
        
        # =========================================================================
        # 🛠️ [사용자 설정 구간] - 물류센터 설정 적용 완료
        # =========================================================================
        
        self.checkbox_name = 'none'
        self.desired_checkboxes = []

        # 오늘 날짜 (숫자만 추출, 예: 28)
        today_day = datetime.now().strftime("%d").lstrip("0") 

        self.dropdown_settings = [
            
            # 1. 날짜 선택 (달력 열기 -> 오늘 날짜 클릭)
            {
                "type": "button", "name": "달력 열기",
                "xpath": "//*[@id='searchForm']/div/div[1]/div[1]/div[2]/div/div[1]/button/div"
            },
            {
                "type": "button", "name": f"오늘 날짜({today_day}일) 선택",
                "xpath": f"//td[contains(text(), '{today_day}')] | //a[contains(text(), '{today_day}')]"
            },

            # 2. 센터 선택 (Custom: 열기 -> 텍스트 클릭)
            {
                "type": "custom", 
                "name": "센터 선택",
                # 버튼 내부의 div를 클릭 (사용자님 원본 경로 복구)
                "open_xpath": "//*[@id='centerIdListContainer']/div/div/button",
                # ⭐️ li 태그 밑에 있는 a 태그의 텍스트를 찾음
                "option_xpath": "//li/a[contains(text(), '{}')]", 
                "value": "INC4" 
            },

            # 3. 캠프 선택 (Select All)
            {
                "type": "button", "name": "캠프 드랍다운 열기",
                "xpath": "//*[@id='campCodeListContainer']/div/div/button/div/div"
            },
            {
                "type": "button", "name": "캠프 Select All 클릭",
                "xpath": "//*[@id='campCodeListContainer']/div/div/div/div[2]/div/button[1]"
            },

            # 4. 정기배송 (Select All)
            {
                "type": "button", "name": "정기배송 드랍다운 열기",
                "xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[2]/div/div/button"
            },
            {
                "type": "button", "name": "정기배송 Select All 클릭",
                "xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div/button[1]"
            },

            # 5. 배송유형 (Select All)
            {
                "type": "button", "name": "배송유형 드랍다운 열기",
                "xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[1]/div/div[1]/button"
            },
            {
                "type": "button", "name": "배송유형 Select All 클릭",
                "xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div/button[1]"
            },

            # 6. ExSD (11시 이후 전부 선택) - ⭐️ 특수 기능
            {
                "type": "time_filter", "name": "ExSD (11시 이후 선택)",
                "open_xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[3]/div/div[1]/button",
                "start_hour": 11
            },

            # 7. 단위 (Parcel 선택)
            {
                "type": "custom", 
                "name": "단위 (Parcel)",
                # 버튼 내부의 div를 클릭 (사용자님 원본 경로 복구)
                "open_xpath": "//*[@id='searchForm']/div/div[1]/div[2]/div[4]/div/div[1]/button",
                # ⭐️ li 태그 밑에 있는 a 태그를 찾음
                "option_xpath": "//li/a[contains(text(), '{}')]",
                "value": "Parcel"
            }
        ]
        # =========================================================================

        settings = self._load_settings() 

        self.user_data_path = tk.StringVar(value=settings.get('user_data_path', r"C:\Users\rmaru\AppData\Local\Google\Chrome\Profile 2"))
        self.profile_dir = tk.StringVar(value=settings.get('profile_dir', "Profile 2"))
        # 실제 사이트 URL로 변경해주세요
        self.target_url = tk.StringVar(value=settings.get('target_url', "https://your-logistics-site.com"))
        
        self.excel_path = tk.StringVar(value=settings.get('excel_path', r"C:\Users\rmaru\OneDrive\바탕 화면\zxc\dsadsa.xlsx")) 
        self.sheet_name = tk.StringVar(value=settings.get('primary_sheet_name', "테스트")) 
        self.start_row = tk.StringVar(value=settings.get('primary_start_row', "34"))
        self.secondary_sheet_name = tk.StringVar(value=settings.get('secondary_sheet_name', "테스트2")) 
        self.secondary_start_row = tk.StringVar(value=settings.get('secondary_start_row', "60"))

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
        
        self.main_button = ttk.Button(button_frame, text="1. 시작하기", style='Blue.TButton', command=self.run_open_browser_and_scrape_thread)
        self.main_button.pack(side='left', fill='x', expand=True, padx=5)
        
        self.quit_button = ttk.Button(button_frame, text="2. 프로그램 종료", style='Red.TButton', command=self.on_closing)
        self.quit_button.pack(side='right', fill='x', expand=True, padx=5)

        self._create_log_section(main_frame) 
        self.update_log("프로그램 준비 완료.", "INFO")

    def _create_setting_section(self, parent, title, fields):
        labelframe = ttk.LabelFrame(parent, text=title, padding="10")
        labelframe.pack(fill='x', padx=5, pady=5)
        for i, (label_text, var) in enumerate(fields):
            ttk.Label(labelframe, text=label_text).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            ttk.Entry(labelframe, textvariable=var, width=60).grid(row=i, column=1, sticky='ew', padx=5, pady=2)

    def _create_excel_section(self, parent):
        labelframe = ttk.LabelFrame(parent, text="엑셀 저장 설정", padding="10")
        labelframe.pack(fill='x', padx=5, pady=5)
        ttk.Label(labelframe, text="Excel File Path:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.excel_path, width=40).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Button(labelframe, text="찾아보기", command=self.browse_excel_path).grid(row=0, column=2, sticky='e', padx=5, pady=2)
        ttk.Label(labelframe, text="[기본] Sheet Name:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.sheet_name, width=15).grid(row=1, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(labelframe, text="[기본] Start Row:").grid(row=1, column=1, sticky='e', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.start_row, width=10).grid(row=1, column=2, sticky='e', padx=5, pady=2)
        ttk.Label(labelframe, text="[보조] Sheet Name:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.secondary_sheet_name, width=15).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        ttk.Label(labelframe, text="[보조] Start Row:").grid(row=2, column=1, sticky='e', padx=5, pady=2)
        ttk.Entry(labelframe, textvariable=self.secondary_start_row, width=10).grid(row=2, column=2, sticky='e', padx=5, pady=2)

    def _create_log_section(self, parent):
        labelframe = ttk.LabelFrame(parent, text="📜 작업 상태 로그", padding="10")
        labelframe.pack(fill='both', expand=True, padx=5, pady=5)
        self.log_text = tk.Text(labelframe, height=12, state='disabled', wrap='word', bg='#1e1e1e', fg='#d4d4d4', font=('Consolas', 10))
        self.log_text.pack(fill='both', expand=True)
        scrollbar = ttk.Scrollbar(labelframe, command=self.log_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.log_text.tag_config("INFO", foreground="#ffffff")
        self.log_text.tag_config("SUCCESS", foreground="#00ff00")
        self.log_text.tag_config("WARNING", foreground="#ffd700") 
        self.log_text.tag_config("ERROR", foreground="#ff5555")
        self.log_text.tag_config("DETAIL", foreground="#87cefa")

    def update_log(self, message, level="INFO"):
        if self.log_text is None: return 
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        full_msg = f"{timestamp} {message}\n"
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, full_msg, level)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.master.update_idletasks()

    def browse_excel_path(self):
        filename = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename: self.excel_path.set(filename)

    def on_closing(self):
        self.update_log("프로그램 종료.", "WARNING")
        if self.driver:
            try: self.driver.quit()
            except: pass
        if self.selection_window: self.selection_window.destroy()
        self.master.destroy()

    def run_open_browser_and_scrape_thread(self):
        self.main_button.config(state='disabled', text="⏳ 작업 진행 중...")
        threading.Thread(target=self._integrated_workflow, daemon=True).start()

    def _integrated_workflow(self):
        self.update_log("--- 작업 시작 ---", "INFO")
        self.open_browser()
        if not self.driver:
            self.main_button.config(state='normal', text="1. 시작하기")
            return

        user_response = messagebox.askokcancel("준비", "로그인 후 원하는 페이지에서 [확인]을 눌러주세요.")
        if not user_response:
            self.main_button.config(state='normal', text="1. 시작하기")
            return

        self._configure_page_settings()
        self.start_scraping()
        self.main_button.config(state='normal', text="1. 시작하기")
    
    def open_browser(self):
        if self.driver:
            self.update_log("🔄 기존 브라우저 재사용", "WARNING")
            try:
                self.driver.get(self.target_url.get())
                return
            except:
                self.driver = None
        
        self.update_log("⏳ 크롬 브라우저 실행...", "WARNING")
        try:
            options = Options()
            options.add_argument(f"user-data-dir={self.user_data_path.get()}") 
            options.add_argument(f"profile-directory={self.profile_dir.get()}") 
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--remote-debugging-port=9222")

            self.driver = webdriver.Chrome(options=options)
            self.driver.get(self.target_url.get())
            self.update_log("✅ 브라우저 접속 성공.", "SUCCESS")
        except Exception as e:
            self.update_log(f"❌ 브라우저 실행 오류: {e}", "ERROR")

    def _restart_scraping(self, current_window):
        if current_window: current_window.destroy()
        self.all_tables = []
        self.current_table_index = 0
        self.update_log("🔄 재탐색 시작", "WARNING")
        self.run_open_browser_and_scrape_thread()

    def _quick_click(self, by_type, xpath_value):
        try:
            element = WebDriverWait(self.driver, 1).until(EC.element_to_be_clickable((by_type, xpath_value)))
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            element.click()
            return True
        except:
            try:
                element = self.driver.find_element(by_type, xpath_value)
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except:
                return False

    def _configure_page_settings(self):
        if not self.driver: return
        self.update_log("⚙️ 페이지 설정 시작...", "WARNING")

        for setting in self.dropdown_settings:
            name = setting.get("name", "Unknown")
            dtype = setting.get("type", "custom")
            try:
                if dtype == "custom":
                    open_xpath = setting.get("open_xpath")
                    option_xpath_fmt = setting.get("option_xpath")
                    value_to_select = setting.get("value")
                    if not self._quick_click(By.XPATH, open_xpath): raise Exception("버튼 없음")
                    final_xpath = option_xpath_fmt.format(value_to_select)
                    if not self._quick_click(By.XPATH, final_xpath): raise Exception("옵션 없음")
                    self.update_log(f"  👉 [Custom] '{name}': {value_to_select} 선택", "DETAIL")

                elif dtype == "button":
                    target_xpath = setting.get("xpath")
                    if self._quick_click(By.XPATH, target_xpath):
                        self.update_log(f"  👉 [Button] '{name}' 클릭 완료", "DETAIL")
                    else:
                        raise Exception("버튼 클릭 실패")

                # ⭐️ [ExSD 전용] 11시 이후 시간 자동 선택
                elif dtype == "time_filter":
                    open_xpath = setting.get("open_xpath")
                    start_hour = setting.get("start_hour", 11)
                    
                    # 1. 드랍다운 열기
                    if not self._quick_click(By.XPATH, open_xpath): raise Exception("드랍다운 열기 실패")
                    time.sleep(0.5) # 목록 로딩 대기

                    # 2. 모든 'a' 태그 가져오기 (시간 목록)
                    # (드랍다운이 열린 상태에서 화면에 보이는 a태그들을 찾습니다)
                    options = self.driver.find_elements(By.TAG_NAME, 'a')
                    selected_count = 0
                    
                    for opt in options:
                        text = opt.text.strip() # 예: "13:00"
                        if ":" in text:
                            try:
                                hour = int(text.split(":")[0]) # "13" -> 13
                                if hour >= start_hour:
                                    # 클릭 시도 (이미 선택된건지 확인 필요하면 class 확인 로직 추가 가능)
                                    self.driver.execute_script("arguments[0].click();", opt)
                                    selected_count += 1
                            except: pass
                    
                    if selected_count > 0:
                        self.update_log(f"  ⏱️ [Time] {start_hour}시 이후 항목 {selected_count}개 선택", "DETAIL")
                    else:
                        self.update_log(f"  ⚠️ [Time] {start_hour}시 이후 항목이 없습니다.", "WARNING")

                time.sleep(0.5) 
            except Exception as e:
                self.update_log(f"⚠️ [패스] '{name}' ({e})", "WARNING")

        self.update_log("✅ 모든 페이지 설정 완료.", "SUCCESS")

    def start_scraping(self):
        self.update_log("⏳ 테이블 탐색 중...", "WARNING")
        time.sleep(1)
        self.all_tables = []
        try:
            try: WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            except: pass
            html_source = self.driver.page_source
            try: self.all_tables = pd.read_html(io.StringIO(html_source))
            except: self.all_tables = []
            num = len(self.all_tables)
            self.update_log(f"✅ 총 {num}개의 테이블 발견.", "SUCCESS")
            if num >= 1: self._open_full_selection_window()
            else: self.update_log("ℹ️ 테이블이 없지만 설정은 완료되었습니다.", "DETAIL")
        except Exception as e:
            self.update_log(f"❌ 탐색 오류: {e}", "ERROR")

    def _finalize_export(self, df_selected: pd.DataFrame, source_window: tk.Toplevel):
        excel_path = self.excel_path.get()
        self.update_log("==========================================", "INFO")
        self.update_log("🚀 엑셀 저장 프로세스 진입", "WARNING")
        df_full = df_selected.replace([float('inf'), float('-inf')], float('nan')).fillna(0)
        existing_sheets = {}
        if os.path.exists(excel_path):
            try: existing_sheets = pd.read_excel(excel_path, sheet_name=None, header=None)
            except: pass
        try:
            self._write_to_excel_file(excel_path, df_full, existing_sheets)
            self.update_log("🎉 저장 완료! (원본 파일 갱신됨)", "SUCCESS")
            source_window.destroy()
        except PermissionError:
            self.update_log("❌ 파일 열림 오류 -> 임시 저장 시도", "ERROR")
            base, ext = os.path.splitext(excel_path)
            temp_path = f"{base}_TEMP_{datetime.now().strftime('%H%M%S')}{ext}"
            try:
                self._write_to_excel_file(temp_path, df_full, existing_sheets)
                messagebox.showinfo("임시 저장", f"파일: {temp_path}\n(원본이 열려있어 임시저장했습니다)")
                self.update_log(f"✅ 임시 저장 완료: {temp_path}", "SUCCESS")
                source_window.destroy()
            except Exception as e:
                self.update_log(f"❌ 임시 저장 실패: {e}", "ERROR")
        except Exception as e:
            self.update_log(f"❌ 저장 실패: {e}", "ERROR")
        self.update_log("==========================================", "INFO")

    def _write_to_excel_file(self, target_path, df_full, existing_sheets):
        USER_SHEET_NAME = self.sheet_name.get() 
        FIXED_SHEET_NAME = self.secondary_sheet_name.get()
        try:
            USER_START_ROW = int(self.start_row.get())
            FIXED_START_ROW = int(self.secondary_start_row.get())
        except: return

        main_current_rows = 0
        if USER_SHEET_NAME in existing_sheets:
            try: main_current_rows = len(existing_sheets[USER_SHEET_NAME])
            except: pass

        fixed_current_rows = 0
        if FIXED_SHEET_NAME in existing_sheets:
            try: fixed_current_rows = len(existing_sheets[FIXED_SHEET_NAME])
            except: pass

        df_sub = pd.DataFrame()
        if USER_SHEET_NAME in existing_sheets:
            try:
                df_test_full = existing_sheets[USER_SHEET_NAME]
                df_sub = df_test_full.iloc[0:32, :].copy() 
            except: pass

        self.update_log("💾 디스크 쓰기 시작...", "WARNING")
        
        with pd.ExcelWriter(target_path, engine='xlsxwriter') as writer:
            wb = writer.book
            fmt = wb.add_format({'border': 0, 'align': 'center', 'valign': 'vcenter'})
            for s_name, data in existing_sheets.items():
                if s_name not in [USER_SHEET_NAME, FIXED_SHEET_NAME]:
                    data.to_excel(writer, sheet_name=s_name, startrow=0, startcol=0, header=False, index=False)
            ws = wb.add_worksheet(USER_SHEET_NAME)
            writer.sheets[USER_SHEET_NAME] = ws
            if USER_SHEET_NAME in existing_sheets:
                existing_sheets[USER_SHEET_NAME].to_excel(writer, sheet_name=USER_SHEET_NAME, startrow=0, startcol=0, header=False, index=False)
            
            main_write_idx = max(USER_START_ROW - 1, main_current_rows)
            self.update_log(f"📍 '{USER_SHEET_NAME}' 저장 위치: {main_write_idx + 1}행", "DETAIL")
            
            for idx, row in df_full.iterrows():
                ws.write_row(main_write_idx + idx, 0, row.tolist(), fmt) 
                
            if not df_sub.empty:
                if FIXED_SHEET_NAME in existing_sheets:
                     existing_sheets[FIXED_SHEET_NAME].to_excel(writer, sheet_name=FIXED_SHEET_NAME, startrow=0, startcol=0, header=False, index=False)
                fixed_write_idx = max(FIXED_START_ROW - 1, fixed_current_rows)
                self.update_log(f"📍 '{FIXED_SHEET_NAME}' 저장 위치: {fixed_write_idx + 1}행", "DETAIL")
                df_sub.to_excel(writer, sheet_name=FIXED_SHEET_NAME, startrow=fixed_write_idx, startcol=0, header=False, index=False)

def _create_dataframe_view(parent_frame, df, height=8):
    tree_frame = ttk.Frame(parent_frame)
    tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
    scroll_y = ttk.Scrollbar(tree_frame)
    scroll_y.pack(side='right', fill='y')
    scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal')
    scroll_x.pack(side='bottom', fill='x')
    columns = list(df.columns)
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=height, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    scroll_y.config(command=tree.yview)
    scroll_x.config(command=tree.xview)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')
    for row in df.itertuples(index=False):
        safe_values = [str(v) for v in row]
        tree.insert("", "end", values=safe_values)
    tree.pack(fill='both', expand=True)
    return tree

def _open_full_selection_window_impl(app):
    if app.selection_window: app.selection_window.destroy()
    win = tk.Toplevel(app.master)
    win.title(f"테이블 선택 (총 {len(app.all_tables)}개 발견)")
    win.geometry("1000x800")
    canvas = tk.Canvas(win)
    scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
    frm = ttk.Frame(canvas)
    frm.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0,0), window=frm, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="top", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    for i, df in enumerate(app.all_tables):
        d = df.dropna(how='all')
        if len(d) < 2: continue 
        lf = ttk.LabelFrame(frm, text=f"📊 Table #{i+1} (크기: {d.shape[0]}행 x {d.shape[1]}열)", padding=10)
        lf.pack(fill='x', padx=10, pady=10)
        btn_frame = ttk.Frame(lf)
        btn_frame.pack(fill='x', pady=(0, 5)) 
        ttk.Button(btn_frame, text="✅ 이 데이터 저장하기", style='Green.TButton', command=lambda d=d: app._finalize_export(d, win)).pack(side='left')
        _create_dataframe_view(lf, d.head(5), height=5)
    bottom_frame = ttk.Frame(win, padding=10)
    bottom_frame.pack(side='bottom', fill='x')
    ttk.Button(bottom_frame, text="🔄 다시 탐색하기", style='Blue.TButton', command=lambda: app._restart_scraping(win)).pack(fill='x')
    app.selection_window = win

def _open_comparison_window_impl(app_instance):
    if app_instance.selection_window: app_instance.selection_window.destroy()
    if app_instance.current_table_index >= len(app_instance.all_tables): return
    df = app_instance.all_tables[app_instance.current_table_index].dropna(how='all')
    win = tk.Toplevel(app_instance.master)
    win.title("테이블 확인")
    win.geometry("900x600")
    ttk.Label(win, text=f"테이블 확인 (#{app_instance.current_table_index + 1})", font=('bold', 12)).pack(pady=10)
    _create_dataframe_view(win, df.head(15), height=15)
    btn_frame = ttk.Frame(win, padding=10)
    btn_frame.pack(fill='x', side='bottom')
    ttk.Button(btn_frame, text="✅ 저장 (이 테이블 맞음)", style='Green.TButton', command=lambda: app._finalize_export(df, win)).pack(side='left', padx=10, expand=True, fill='x')
    ttk.Button(btn_frame, text="⏭️ 다음 테이블 보기", style='Blue.TButton', command=lambda: _move_to_next_table_impl(app_instance, win)).pack(side='right', padx=10, expand=True, fill='x')
    app_instance.selection_window = win

def _move_to_next_table_impl(app_instance, current_window):
    if current_window: current_window.destroy()
    app_instance.current_table_index += 1
    app_instance.master.after(10, lambda: _open_comparison_window_impl(app_instance)) 

if __name__ == "__main__":
    root = tk.Tk()
    app = WebScraperApp(root)
    WebScraperApp._open_comparison_window = _open_comparison_window_impl
    WebScraperApp._move_to_next_table = _move_to_next_table_impl
    WebScraperApp._open_full_selection_window = _open_full_selection_window_impl
    root.mainloop()