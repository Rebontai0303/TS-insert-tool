'''
author rebontai 20251208
============================
ZSDT8004 主檔及附檔輸入程式
============================ 
20251231 新增派車資料上傳SAP_LOGISTICS_TP功能  
'''
from pyodbc import connect, IntegrityError
from pandas import DataFrame, read_excel, ExcelFile
from tkinter import messagebox, filedialog, StringVar, ttk
from time import sleep
from threading import Thread
from sys import exit, getwindowsversion
from pywinstyles import change_header_color, apply_style
from datetime import datetime
from socket import gethostname, gethostbyname, gaierror
from os import getlogin
from requests import post
import sv_ttk
import darkdetect
import tkinter as tk

def print_message(message, function_name, type):
    '''
    print message.
    顯示通知訊息之tk視窗
    種類涵蓋
    err: 錯誤訊息
    info: 一般訊息
    war: 警告訊息

    Args:
        message (str): 訊息內容
        function_name (str): 函式名稱
        type (str): 訊息類型
    Return:
        NA.    
    '''
    match type:
        case "err":
            print (f"❗{function_name}發生錯誤: {message}❗")
            log_text.insert("end", f"❗{function_name}發生錯誤: {message}❗\n")
            log_text.see("end")
            pass_error(function_name, message)
            sleep(10)
        case "info":
            print (f"ℹ️{function_name}訊息: {message}ℹ️")
            log_text.insert("end", f"ℹ️{function_name}訊息: {message}ℹ️\n")
            log_text.see("end")
        case "war":
            print (f"⚠️{function_name}警告: {message}⚠️")
            log_text.insert("end", f"⚠️{function_name}警告: {message}⚠️\n")
            log_text.see("end")  
  

def sql_connect()->object:
    '''
    連線到MSSQL資料庫

    Args:
        table (str): 資料庫名稱
    Returns:
        cursor, conn (object): 資料庫連線物件
    '''
    server   = '192.168.16.109'
    username = 'sqluser'
    password = 'pintai2011'
    database = 'sapedi_test'
    conn = connect('DRIVER={ODBC Driver 11 for SQL Server};'
                   f'SERVER={server};'
                   f'DATABASE={database};'
                   f'UID={username};'
                   f'PWD={password}')
    cursor = conn.cursor()
    return cursor, conn

def check_format(config):
    '''
    格式檢查共用函式

    Args:
        config (dict): 檢查參數
            Path (str): Excel檔案路徑
            Table (str): 資料庫表格名稱
            Class_name (str): 呼叫此函式的類別名稱
    Returns:
        df (DataFrame): 回傳符合格式的DataFrame物件
        db_columns_map.keys() (list): 回傳資料庫欄位名稱列表
    '''
    def _message(text, type):
        print_message(text, f'{class_name}.{table}.check_format', type)
    
    class_name = config['class_name']
    path       = config['path']
    table      = config['table']

    _message('開始檢核作業', 'info')
    # 讀取Excel檔案
    xlsx = ExcelFile(path)
    _message(f'Excel檔案包含的工作表: {xlsx.sheet_names}', 'info')

    # ============================
    # 工作表名稱檢核
    # ============================
    # 檢查工作表是否符合指定名稱需求
    _message(f'尋找{table}中...🔍', 'info')
    if table in xlsx.sheet_names:
        _message(f'工作表名稱符合要求: {table}', 'info')
    else:
        _message(f'未找尋到工作表, 終止程式', 'err')
        exit(0)
    
    # 抓取資料庫欄位格式進行檢核
    _message(f'欄位檢核中...🔍', 'info')
    cursor, _ = sql_connect()
    cursor.execute(f'SELECT * FROM {table} WHERE 1=0')
    db_columns_map = {
        col[0]:{
        'data_type':col[1],
        'size':col[3]
        }
        for col in cursor.description
    }
    # 先刪除GUID, 後補上
    col_dict = db_columns_map.copy()
    del col_dict['GUID']
    # ============================
    # 欄位數量檢核
    # ============================
    # 檢查表格中的欄位數量是否符合指定需求
    df = read_excel(path, sheet_name = table, keep_default_na=False)
    # 核對
    if len(df.columns) != len(col_dict):
        _message(f'欄位數量不符合要求, 終止程式', 'err')
        exit(0)
    else:
        _message(f'欄位數量符合要求', 'info')

    # 重新命名欄位名稱
    db_cols = list(col_dict.keys())
    df.columns = db_cols
    # ============================
    # 欄位資料長度檢核
    # ============================
    for col in col_dict:
        max_size = col_dict[col]['size']
        for index, value in df[col].items():
            if isinstance(value, str) and len(value) > max_size:
                _message(f'欄位 {col} 第 {index+2} 列資料長度超過限制 ({len(value)} > {max_size}), 終止程式', 'err')
                exit(0)
            else:
                continue
    
    # 讀取         
    _message(f'檢核完成, 讀取到 {len(df)} 筆資料', 'info')
    return df, db_columns_map.keys()

def get_contril_file(config):
    '''
    取得控制檔插入SQL語法

    Args:
        config (dict): 插入參數包
            spec (str): Tcode對應碼
            table (str): Table名稱
            data_count (int): 資料量
            table_count (int): 檔案數量
    Returns:
        SPEC_ID          : Tcode對應碼
        GUID             : SAP識別碼
        sender           : 資料來源方
        receiver         : 資料接收方
        Table_Name       : Table名稱
        Data_Count       : 資料量
        Table_Count      : 檔案數量
        Read_Flag        : SAP讀取註記
        Sender_Datetime  : 轉入日期
        Receiver         : SAP讀取時間
    '''
    spec = config['spec']
    guid = config['guid']
    sender = 'Python'
    receiver = 'SAP'
    table_name = config['table']
    data_count = config['data_count']
    table_count = config['table_count']
    read_flag = ''
    sender_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:23]
    receiver_datetime = None
    sql = f'INSERT INTO Control_File (SPEC_ID, GUID, SENDER, RECEIVER, TABLE_NAME, DATA_COUNT, TABLE_COUNT, READ_FLAG, SENDER_DATETIME, RECEIVER_DATETIME) VALUES (?,?,?,?,?,?,?,?,?,?)'
    return sql, [spec, guid, sender, receiver, table_name, data_count, table_count, read_flag, sender_datetime, receiver_datetime]

def guid_():
    '''
    產生GUID

    Returns:
        guid (str): 回傳GUID字串
    '''
    global num
    date = datetime.now().strftime("%Y%m%d%H%M%S")
    num = "1"
    guid_num = str(num).zfill(6)
    guid = date + guid_num
    return guid

def get_system_info():
    '''
    獲取系統資訊(IP及Windows帳號)
    Returns:
        info (dict): 回傳系統資訊字典
    '''
    info = {}

    # 獲取IP
    try:
        hostname = gethostname()
        info['ip'] = gethostbyname(hostname)
    except gaierror:
        info['ip'] = "無法獲取本機 IP 位址"

    # 獲取Windows
    try:
        info['username'] = getlogin()
    except OSError:
        info['username'] = "無法獲取 Windows 帳號"
    return info


def pass_error(part, e):
    info = get_system_info()
    webhook = f'https://discord.com/api/webhooks/1409531559358365839/uXpQJl_JZbOVVZlCyMDdxy_eTNHWExTNLCR_gMPeg0m6qGOMz0t_TQaYHXHeD-k2ZYMP'
    url = webhook
    message = f"{part} 程式區段出現錯誤 : {e}, ip : {info.get('ip')}, username : {info.get('username')}"    
    payload = {"content":message}
    headers = {"Content-Type" : "application/json"}
    post(url, json=payload, headers=headers)

# ============================ZSDT8004============================

class ZSDT8004:
    def __init__(self):
        self.table   = 'ZSDT8004'
        self.table_A = 'ZSDT8004A'

    def execute_task(self):
        '''
        執行ZSDT8004轉檔作業

        Args:
            None
        Returns:
            None
        '''
        print_message('執行ZSDT8004轉檔作業', 'ZSDT8004.execute_task', 'info')
        file_path = filedialog.askopenfilename(title="選擇回單Excel檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            print_message('未選擇檔案，作業取消', 'ZSDT8004.execute_task', 'war')
            exit(0)
        print_message(f'選擇的檔案路徑: {file_path}', 'ZSDT8004.execute_task', 'info')
        print_message('開始讀取Excel檔案並寫入資料庫', 'ZSDT8004.execute_task', 'info') 

        # 定義檢核參數
        config = {
            'path':file_path,
            'table':self.table,
            'class_name':self.__class__.__name__
        }
        # 檢核
        df_zsdt8004, col_zsdt8004 = check_format(config)

        config.clear()

        # 定義檢核參數
        config = {
            'path':file_path,
            'table':self.table_A,
            'class_name':self.__class__.__name__
        }
        # 檢核
        df_zsdt8004a, col_zsdt8004a = check_format(config)

        config.clear()

        # 連線到資料庫
        try:
            cursor, conn = sql_connect()
            print_message('資料庫連線成功', 'ZSDT8004.execute_task', 'info')
        except Exception as e:
            print_message(f'資料庫連線失敗: {e}', 'ZSDT8004.sql_connect', 'err')
            exit(0)

        # 定義control file參數
        guid = guid_()
        config = {
            'guid':guid,
            'spec':'SD-P31',
            'table':self.table,
            'data_count':len(df_zsdt8004),
            'table_count':2
        }
        
        # 插入資料到ZSDT8004
        try:
            print_message('插入資料到ZSDT8004中', 'ZSDT8004.insert_data', 'info')
            for row in df_zsdt8004.itertuples():
                row = [guid] + list(row)[1:]
                cursor.execute(f"INSERT INTO {self.table} ({",".join(col_zsdt8004)}) VALUES ({",".join("?"*len(col_zsdt8004))})", row)
            sql, ctrl_row = get_contril_file(config)
            cursor.execute(sql, ctrl_row)
            conn.commit()
            print_message('ZSDT8004資料插入完成', 'ZSDT8004.insert_data', 'info')
        except IntegrityError as e:
            print_message(f'重複插入資料: {e}', 'ZSDT8004.ZSDT8004.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)
        except Exception as e:
            print_message(f'發生預期外錯誤: {e}', 'ZSDT8004.ZSDT8004.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)

        # 清除參數
        config.clear()

        # 定義control file參數
        guid_a = guid_()
        config = {
            'guid':guid_a,
            'spec':'SD-P31',
            'table':self.table_A,
            'data_count':len(df_zsdt8004a),
            'table_count':2
        }
        
        # 插入資料到ZSDT8004A
        try:
            print_message('插入資料到ZSDT8004A中', 'ZSDT8004.insert_data', 'info')
            for row in df_zsdt8004a.itertuples():
                row = [guid_a] + list(row)[1:]
                cursor.execute(f"INSERT INTO {self.table_A} ({",".join(col_zsdt8004a)}) VALUES ({",".join("?"*len(col_zsdt8004a))})", row)
            sql_a, ctrl_row_a = get_contril_file(config)
            cursor.execute(sql_a, ctrl_row_a)
            conn.commit()
            print_message('ZSDT8004A資料插入完成', 'ZSDT8004.insert_data', 'info')
        except IntegrityError as e:
            print_message(f'重複插入資料: {e}', 'ZSDT8004.ZSDT8004A.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)
        except Exception as e:
            print_message(f'發生預期外錯誤: {e}', 'ZSDT8004.ZSDT8004A.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)

        # 清除參數
        config.clear()

        # Close connection
        conn.close()

# ===============================TP==================================
        
class TP():
    def __init__(self):
        self.table = 'SAP_LOGISTICS_TP'

    def execute_task(self):
        '''
        執行TP轉檔作業

        Args:
            None
        Returns:
            None
        '''
        print_message('執行TP轉檔作業', 'TP.execute_task', 'info')
        file_path = filedialog.askopenfilename(title="選擇派車Excel檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            print_message('未選擇檔案，作業取消', 'TP.execute_task', 'war')
            exit(0)
        print_message(f'選擇的檔案路徑: {file_path}', 'TP.execute_task', 'info')
        print_message('開始讀取Excel檔案並寫入資料庫', 'TP.execute_task', 'info') 

        # 定義檢核參數
        config = {
            'path':file_path,
            'table':self.table,
            'class_name':self.__class__.__name__
        }
        # 檢核
        df_tp, col_tp = check_format(config)

        config.clear()

        # 連線到資料庫
        try:
            cursor, conn = sql_connect()
            print_message('資料庫連線成功', 'TP.execute_task', 'info')
        except Exception as e:
            print_message(f'資料庫連線失敗: {e}', 'TP.sql_connect', 'err')
            exit(0)

        # 定義control file參數
        guid = guid_()
        config = {
            'guid':guid,
            'spec':'SD-TP',
            'table':self.table,
            'data_count':len(df_tp),
            'table_count':1
        }

        # 插入資料到SAP_LOGISTICS_TP
        try:
            print_message('插入資料到SAP_LOGISTICS_TP中', 'TP.insert_data', 'info')
            for row in df_tp.itertuples():
                row = [guid] + list(row)[1:]
                cursor.execute(f"INSERT INTO {self.table} ({','.join(col_tp)}) VALUES ({','.join('?'*len(col_tp))})", row)
            sql, ctrl_row = get_contril_file(config)
            cursor.execute(sql, ctrl_row)
            conn.commit()
            print_message('派車資料插入完成', 'TP.insert_data', 'info')
        except IntegrityError as e:
            print_message(f'重複插入資料: {e}', 'TP.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)
        except Exception as e:
            print_message(f'發生預期外錯誤: {e}', 'TP.insert_data', 'err')
            conn.rollback()
            conn.close()
            exit(0)

# ===========================Main Program============================ 

def main():
    def execute():
        '''
        依據Radiobutton判定執行作業

        Args:
            None
        Returns:
            None
        '''
        if type:
            match type:
                case 'SD-P31':
                    task = ZSDT8004()
                    task.execute_task()
                case 'SD-TP':
                    task = TP()
                    task.execute_task()    
        else:
            print("請選擇一個選項！")
            return
        
    def open_main_window():
        '''
        setting tkinter main window.
        主要視窗設定
        使用主題 - darkdetect
        使用字體 - 標楷體(DFKai-SB)

        Args:
            NA.
        Return:
            NA.
        '''
        global root, theme_switch, option
        # tkinter視窗設定
        root = tk.Tk()
        # 主題參數
        sv_ttk.set_theme(darkdetect.theme())
        # 視窗標題
        root.title('轉檔程式')
        # 視窗外觀
        window_width = root.winfo_screenwidth()    # 取得螢幕寬度
        window_height = root.winfo_screenheight()  # 取得螢幕高度
        width = 400
        height = 150
        left = int((window_width - width)/2)       # 計算左上 x 座標
        top = int((window_height - height)/2)      # 計算左上 y 座標
        root.geometry(f"{width}x{height}+{left}+{top}")
        root.resizable(False, False)               # 設定視窗不可調整大小
        
        # tk視窗設定
        option = StringVar()
        # option.set('SD-P31') # 預設第一個為選項 

        # Radiobutton選項
        rd1 = ttk.Radiobutton(
            root, 
            text = 'SD-P31 回單資料上傳', 
            variable = option, 
            value = 'SD-P31'
            # font=("DFKai-SB", 12)
        )
        rd1.pack()
        rd2 = ttk.Radiobutton(
            root,
            text = 'SD-TP 派車資料上傳', 
            variable = option, 
            value = 'SD-TP'
            # font=("DFKai-SB", 12)
        )
        rd2.pack()
        # Radiobutton(
        #     root, 
        #     text = '轉檔名稱', 
        #     variable = option, 
        #     value = '模組-編碼'
        # ).pack()
        # 執行按鈕
        ttk.Button(
            root, 
            text="執行", 
            command=on_start,
        ).pack()

        # 淺深滑桿設定
        frame = ttk.Frame(root, padding="10").pack(expand=True, fill="both")
        bottom_frame = ttk.Frame(frame)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        #copyright
        ttk.Label(bottom_frame, text="Copyright © 2025 Rebontai", font=("Arial", 10)).pack(side=tk.LEFT, padx=10, pady=10)  

        # 元件style設定
        style = ttk.Style()
        style.configure(
            "TButton", 
            font=("DFKai-SB", 15)
        )
        style.configure(
            "Switch.TCheckbutton", 
            font=("DFKai-SB", 10)
        )  
        style.configure(
            "TRadiobutton", 
            font=("DFKai-SB", 12)
        )
        # theme_switch = ttk.Checkbutton(
        #     bottom_frame, 
        #     style="Switch.TCheckbutton"
        # )
        
        
        # 滑桿綁定當前主題
        # theme_switch.pack(side=tk.RIGHT, padx=10, pady=10)
        # 根據主題設定font和label
        # if sv_ttk.get_theme() == "dark":
        #     theme_switch.configure(text="深色模式")
        #     theme_switch.state(["selected"])
        # else:
        #     theme_switch.configure(text="淺色模式")
        #     theme_switch.state(["!selected"])  
        
        # # 設定主題樣式
        # apply_titlebar_theme()

        # 關閉視窗即終止
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()

    def on_start():
        '''
        when the "Start" button is clicked, do.
        當按下開始按鈕後檢查欄位是否為null
        若為null則跳出警告視窗
        若不為null則隱藏主視窗並開啟日誌視窗

        Args:
            NA.
        Return:
            NA.
        '''
        global type
        type = option.get().strip()
        if not type:
            messagebox.showwarning("警告", "❗請選擇要執行的選項❗")
            return
        root.withdraw()  # 隱藏主視窗
        open_log_window()  # 開啟日誌視窗        
        
    def open_log_window():
        '''
        open a new window to display logs.
        開啟新視窗顯示程式運作訊息
        使用多執行緒(Thread)方式載入後續動作, 避免tkinter無回應

        Args:
            NA.
        Return:
            NA.    
        '''    
        global log_text
        
        # tk視窗設定
        log_window = tk.Toplevel()
        log_window.title("轉檔程式")
        window_width = log_window.winfo_screenwidth()    # 取得螢幕寬度
        window_height = log_window.winfo_screenheight()  # 取得螢幕高度
        width = 800
        height = 320
        left = int((window_width - width)/2)       # 計算左上 x 座標
        top = int((window_height - height)/2)      # 計算左上 y 座標
        log_window.geometry(f"{width}x{height}+{left}+{top}")

        # title
        tk.Label(log_window, text="⏳運行中...", font = ("DFKai-SB", 11)).pack(side=tk.TOP, anchor=tk.NW)
        text_frame = ttk.Frame(log_window)
        text_frame.pack(fill="both", expand=True)

        # log window
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        # log text config
        log_text = tk.Text(text_frame, wrap="word", font=("DFKai-SB", 11), yscrollcommand=scrollbar.set)
        log_text.pack(fill="both", expand=True)
        scrollbar.config(command=log_text.yview)

        # thread載入後續動作, 避免tkinter無回應
        Thread(target=execute, daemon=True).start()

        # 關閉視窗即終止
        log_window.protocol("WM_DELETE_WINDOW", on_closing)

    def on_closing():
        '''
        when the window is closed, and it terminates the entire program.
        當視窗關閉時, 終止整個程式

        Args:
            NA.
        Return:
            NA.
        '''
        root.destroy()
        exit(0)

    def apply_titlebar_theme():
        '''
        applies the theme to the window's title bar on supported Windows versions.
        根據Windows版本設定視窗標題列樣式

        Args:
            NA.
        Return:
            NA.
        '''
        version = getwindowsversion()
        if version.major == 10 and version.build >= 22000:
            change_header_color(root, "#1c1c1c" if sv_ttk.get_theme() == "dark" else "#fafafa")
        elif version.major == 10:
            apply_style(root, "dark" if sv_ttk.get_theme() == "dark" else "normal")
            root.wm_attributes("-alpha", 0.99)
            root.wm_attributes("-alpha", 1)              
    # Main program starts here
    open_main_window()

if __name__ == "__main__":
    main()