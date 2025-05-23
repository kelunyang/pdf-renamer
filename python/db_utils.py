import sqlite3
import time
import os
import gc
import threading
import queue
import csv
from contextlib import contextmanager

# 全局變量
db_lock = threading.Lock()
# 使用線程本地存儲來確保每個線程使用自己的連接
thread_local = threading.local()
connection_pool = queue.Queue(maxsize=10)  # 最多10個連接
db_file = 'pdf_processing.db'

# 數據庫連接管理
@contextmanager
def db_connection():
    """
    獲取數據庫連接的上下文管理器（線程安全）
    使用線程本地存儲確保每個線程使用自己的連接
    
    返回:
        sqlite3.Connection: 數據庫連接
    """
    # 獲取當前線程ID
    current_thread_id = threading.get_ident()
    
    # 創建新連接（每次調用都創建新連接，避免跨線程共享）
    conn = sqlite3.connect(db_file, timeout=30)
    conn.execute('PRAGMA journal_mode=WAL')  # 使用WAL模式提高並發性能
    
    try:
        # 返回連接
        yield conn
        
        # 提交事務
        conn.commit()
    except sqlite3.Error as e:
        # 發生錯誤時回滾事務
        if conn:
            conn.rollback()
        raise e
    finally:
        # 確保連接被關閉
        if conn:
            conn.close()

# 數據庫初始化函數
def init_database():
    """
    初始化數據庫（線程安全）
    創建必要的表格結構
    """
    with db_lock:
        with db_connection() as conn:
            # 創建文件表
            conn.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    path TEXT PRIMARY KEY,
                    status INTEGER,
                    message TEXT,
                    thread INTEGER,
                    start_time REAL,
                    end_time REAL
                )
            ''')

def cleanup_database():
    """關閉所有數據庫連接並刪除數據庫文件（線程安全）
    
    此函數專注於清理數據庫相關資源：
    1. 關閉所有活躍的數據庫連接
    2. 導出數據庫內容到CSV
    3. 刪除臨時數據庫文件
    4. 清空SQLite緩存
    """
    global connection_pool
    
    # 首先導出數據庫內容到CSV
    try:
        export_database_to_csv()
    except Exception as e:
        print(f"導出數據庫到CSV時出錯: {e}")
        from log_utils import log_message
        log_message(f"導出數據庫到CSV時出錯: {e}", level='警告')
    
    with db_lock:
        # 清空連接池並關閉所有連接
        while not connection_pool.empty():
            try:
                conn = connection_pool.get(block=False)
                if conn:
                    conn.close()
            except queue.Empty:
                break
        
        # 獲取當前線程ID
        current_thread_id = threading.get_ident()
        
        # 獲取並關閉所有活躍連接，但只關閉當前線程創建的連接
        connections = [obj for obj in gc.get_objects() if isinstance(obj, sqlite3.Connection)]
        for conn in connections:
            try:
                # 檢查連接是否由當前線程創建
                # 注意：這裡我們不能直接檢查，因為SQLite連接沒有公開線程ID
                # 所以我們使用try-except來安全地嘗試關閉
                if conn:
                    conn.close()
            except Exception as e:
                # 如果出現跨線程錯誤，只是記錄而不拋出異常
                print(f"關閉數據庫連接時出錯: {e}")
                from log_utils import log_message
                log_message(f"關閉數據庫連接時出錯: {e}", level='警告')
        
        # 刪除數據庫文件
        if os.path.exists(db_file):
            try:
                for _ in range(5):  # 最多嘗試5次
                    try:
                        os.remove(db_file)
                        print(f"已刪除臨時數據庫文件: {db_file}")
                        from log_utils import log_message
                        log_message("已刪除臨時數據庫文件", level='信息')
                        break
                    except PermissionError:
                        print(f"文件被佔用，等待重試... (剩餘嘗試次數: {4-_})")
                        time.sleep(1)
            except Exception as e:
                print(f"刪除數據庫文件失敗: {e}")
                from log_utils import log_message
                log_message(f"刪除數據庫文件失敗: {e}", level='错误')

        # 清空SQLite緩存
        if hasattr(sqlite3, '_sqlite3'):
            sqlite3._sqlite3 = None

def get_file_status(file_path):
    """
    獲取文件處理狀態
    
    參數:
        file_path (str): 文件路徑
        
    返回:
        dict or None: 文件狀態信息
    """
    with db_connection() as conn:
        cursor = conn.execute(
            "SELECT path, status, message, thread, start_time, end_time FROM files WHERE path = ?",
            (file_path,)
        )
        result = cursor.fetchone()
    
    if result:
        return {
            'path': result[0],
            'status': result[1],
            'message': result[2],
            'thread': result[3],
            'start_time': result[4],
            'end_time': result[5]
        }
    else:
        return None

def get_all_file_statuses():
    """
    獲取所有文件處理狀態
    
    返回:
        list: 文件狀態信息列表
    """
    with db_connection() as conn:
        cursor = conn.execute(
            "SELECT path, status, message, thread, start_time, end_time FROM files"
        )
        results = cursor.fetchall()
    
    file_statuses = []
    for result in results:
        file_statuses.append({
            'path': result[0],
            'status': result[1],
            'message': result[2],
            'thread': result[3],
            'start_time': result[4],
            'end_time': result[5]
        })
    
    return file_statuses

def get_pending_files(limit=None):
    """
    獲取待處理的文件列表
    
    參數:
        limit (int): 限制返回的文件數量，默認為None（返回所有）
        
    返回:
        list: 待處理的文件路徑列表
    """
    query = "SELECT path FROM files WHERE status = 0"
    params = ()
    
    if limit is not None:
        query += " LIMIT ?"
        params = (limit,)
    
    with db_connection() as conn:
        cursor = conn.execute(query, params)
        results = cursor.fetchall()
    
    return [result[0] for result in results]

def update_file_status(file_path, status, message=None, thread_id=None, start_time=None, end_time=None):
    """
    更新文件處理狀態
    
    參數:
        file_path (str): 文件路徑
        status (int): 狀態碼 (0: 等待, 1: 處理中, 2: 完成, 3: 失敗)
        message (str): 狀態消息
        thread_id (int): 線程ID
        start_time (float): 開始處理時間（Unix時間戳）
        end_time (float): 結束處理時間（Unix時間戳）
    """
    try:
        # 獲取當前文件狀態
        current_status = get_file_status(file_path)
        
        # 如果文件不存在於數據庫中，則創建一個新的狀態記錄
        if current_status is None:
            current_status = {
                'path': file_path,
                'status': 0,
                'message': '',
                'thread': None,
                'start_time': None,
                'end_time': None
            }
        
        # 更新狀態
        if status is not None:
            current_status['status'] = status
        
        # 更新消息
        if message is not None:
            current_status['message'] = message
        
        # 更新線程ID
        if thread_id is not None:
            current_status['thread'] = thread_id
        else:
            # 使用當前線程ID
            current_status['thread'] = threading.get_ident()
        
        # 更新時間戳
        if start_time is not None:
            current_status['start_time'] = start_time
        elif status == 1 and current_status['start_time'] is None:  # 開始處理
            current_status['start_time'] = time.time()
        
        if end_time is not None:
            current_status['end_time'] = end_time
        elif status in [2, 3] and current_status['end_time'] is None:  # 完成或失敗
            current_status['end_time'] = time.time()
        
        # 更新數據庫
        with db_connection() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO files (path, status, message, thread, start_time, end_time) VALUES (?, ?, ?, ?, ?, ?)",
                (file_path, current_status['status'], current_status['message'], 
                 current_status['thread'], current_status['start_time'], current_status['end_time'])
            )
    except Exception as e:
        from log_utils import log_message
        log_message(f"更新數據庫狀態時出錯: {e}", level='警告')
        print(f"更新數據庫狀態時出錯: {e}")

def export_database_to_csv():
    """
    將數據庫內容導出為CSV文件（帶BOM標記以支持中文Excel正確顯示）
    
    返回:
        bool: 是否成功導出
    """
    try:
        # 創建輸出目錄
        output_dir = "."
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 導出文件狀態表
        csv_path = os.path.join(output_dir, f"files_status_{timestamp}.csv")
        
        # 獲取所有文件狀態
        with db_connection() as conn:
            cursor = conn.execute(
                "SELECT path, status, message, thread, start_time, end_time FROM files"
            )
            results = cursor.fetchall()
        
        # 寫入CSV文件（帶BOM標記）
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            # 寫入標題行
            writer.writerow(["文件路徑", "狀態", "消息", "線程ID", "開始時間", "結束時間"])
            # 寫入數據行
            for result in results:
                # 將時間戳轉換為可讀格式
                start_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(result[4])) if result[4] else ""
                end_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(result[5])) if result[5] else ""
                
                # 將狀態碼轉換為文字
                status_text = {
                    0: "等待處理",
                    1: "處理中",
                    2: "已完成",
                    3: "失敗"
                }.get(result[1], str(result[1]))
                
                writer.writerow([result[0], status_text, result[2], result[3], start_time, end_time])
        
        print(f"已將文件狀態導出到 {csv_path}")
        from log_utils import log_message
        log_message(f"已將文件狀態導出到 {csv_path}", level='信息')
        return True
    except Exception as e:
        print(f"導出數據庫到CSV時出錯: {e}")
        from log_utils import log_message
        log_message(f"導出數據庫到CSV時出錯: {e}", level='错误')
        return False

def add_files_to_database(file_list):
    """
    將文件添加到數據庫中
    
    參數:
        file_list (list): 文件路徑列表
        
    返回:
        int: 添加的文件數量
    """
    added_count = 0
    
    try:
        with db_connection() as conn:
            for file_path in file_list:
                # 檢查文件是否已存在於數據庫中
                cursor = conn.execute("SELECT 1 FROM files WHERE path = ?", (file_path,))
                exists = cursor.fetchone() is not None
                
                if not exists:
                    # 插入新記錄
                    conn.execute(
                        "INSERT INTO files (path, status, message, thread, start_time, end_time) VALUES (?, ?, ?, ?, ?, ?)",
                        (file_path, 0, "等待處理", None, None, None)
                    )
                    added_count += 1
    except Exception as e:
        from log_utils import log_message
        log_message(f"添加文件到數據庫時出錯: {e}", level='警告')
        print(f"添加文件到數據庫時出錯: {e}")
    
    return added_count

def add_files_to_database(file_list):
    """
    將文件列表添加到數據庫中
    
    參數:
        file_list (list): 文件路徑列表
        
    返回:
        int: 添加的文件數量
    """
    added_count = 0
    
    with db_connection() as conn:
        for file_path in file_list:
            try:
                # 檢查文件是否已存在於數據庫中
                cursor = conn.execute("SELECT 1 FROM files WHERE path = ?", (file_path,))
                if cursor.fetchone() is None:
                    # 文件不存在，添加到數據庫
                    conn.execute(
                        "INSERT INTO files (path, status, message, thread, start_time, end_time) VALUES (?, ?, ?, ?, ?, ?)",
                        (file_path, 0, "等待處理", None, None, None)
                    )
                    added_count += 1
            except Exception as e:
                print(f"添加文件到數據庫時出錯: {e}")
                try:
                    from log_utils import log_message
                    log_message(f"添加文件到數據庫時出錯: {e}", level='警告')
                except ImportError:
                    pass
    
    return added_count