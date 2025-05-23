import os
import threading
import time
import shutil
import sqlite3

from db_utils import db_connection
from log_utils import log_message
from worker_utils import get_worker_status, wait_for_ui_update

# 全局變量
ui_thread_running = False
ui_thread_handle = None

def get_terminal_size():
    """獲取終端大小"""
    try:
        columns, lines = shutil.get_terminal_size()
        return type('TerminalSize', (), {'columns': columns, 'lines': lines})
    except Exception:
        # 如果無法獲取終端大小，使用默認值
        return type('TerminalSize', (), {'columns': 80, 'lines': 24})

def clear_screen():
    """清除屏幕"""
    os.system('cls' if os.name == 'nt' else 'clear')

def display_files_status():
    """顯示文件處理狀態"""
    clear_screen()
    print("PDF檔案處理狀態:\n")
    
    # 獲取終端大小
    terminal_size = get_terminal_size()
    terminal_width = terminal_size.columns
    
    # 計算每行可以顯示的文件數量
    files_per_line = max(1, terminal_width // 30)  # 假設每個文件名平均30個字符
    
    try:
        # 按狀態分類文件
        pending_files = []
        processing_files = []
        failed_files = []
        completed_files = []
        
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT path, status, message FROM files')
            for row in cursor.fetchall():
                path, status, message = row
                file_name = os.path.basename(path)
                if status == 0:  # 待處理
                    pending_files.append(f"\033[90m{file_name}\033[0m")  # 深灰色
                elif status == 1:  # 處理中
                    processing_files.append(f"\033[93m{file_name}\033[0m")  # 黃色
                elif status == 2:  # 已處理成功
                    completed_files.append(f"\033[92m{file_name}\033[0m")  # 綠色
                elif status == 3:  # 處理失敗
                    failed_files.append(f"\033[91m{file_name} ({message})\033[0m")  # 紅色
        
        # 計算統計信息
        pending = len(pending_files)
        processing = len(processing_files)
        failed = len(failed_files)
        completed = len(completed_files)
        total = pending + processing + completed + failed
        
        # 計算可用的顯示行數 (保留一些行給統計信息和進度條)
        available_lines = terminal_size.lines - 15  # 預留行給標題、統計和進度條
        max_files_per_category = available_lines // 3  # 每個類別最多顯示的行數
        
        # 顯示處理中的文件 (優先顯示)
        if processing_files:
            print("處理中:")
            for i in range(0, min(len(processing_files), max_files_per_category * files_per_line), files_per_line):
                end_idx = min(i + files_per_line, len(processing_files))
                print(" ".join(processing_files[i:end_idx]))
            if len(processing_files) > max_files_per_category * files_per_line:
                print(f"...等 {len(processing_files) - max_files_per_category * files_per_line} 個處理中的檔案")
            print()
        
        # 顯示失敗的文件
        if failed_files:
            print("處理失敗:")
            for i in range(0, min(len(failed_files), max_files_per_category * files_per_line), files_per_line):
                end_idx = min(i + files_per_line, len(failed_files))
                print(" ".join(failed_files[i:end_idx]))
            if len(failed_files) > max_files_per_category * files_per_line:
                print(f"...等 {len(failed_files) - max_files_per_category * files_per_line} 個失敗的檔案")
            print()
        
        # 顯示待處理的文件
        if pending_files:
            print("待處理:")
            for i in range(0, min(len(pending_files), max_files_per_category * files_per_line), files_per_line):
                end_idx = min(i + files_per_line, len(pending_files))
                print(" ".join(pending_files[i:end_idx]))
            if len(pending_files) > max_files_per_category * files_per_line:
                print(f"...等 {len(pending_files) - max_files_per_category * files_per_line} 個待處理的檔案")
            print()
        
        # 顯示已完成的文件（如果有空間）
        if completed_files and available_lines > 0:
            print("已完成:")
            for i in range(0, min(len(completed_files), max_files_per_category * files_per_line), files_per_line):
                end_idx = min(i + files_per_line, len(completed_files))
                print(" ".join(completed_files[i:end_idx]))
            if len(completed_files) > max_files_per_category * files_per_line:
                print(f"...等 {len(completed_files) - max_files_per_category * files_per_line} 個已完成的檔案")
            print()
        
        print(f"統計: 總計 {total} 個檔案, 已完成 {completed}, 處理中 {processing}, 待處理 {pending}, 失敗 {failed}\n")
        
        # 如果有進度，顯示進度條
        if total > 0:
            progress = completed / total
            bar_width = terminal_width - 20
            filled_width = int(bar_width * progress)
            bar = f"[{'#' * filled_width}{'-' * (bar_width - filled_width)}] {progress:.1%}"
            print(f"\n{bar}\n")
    
    except Exception as e:
        log_message(f"顯示檔案狀態時出錯: {e}", level='警告')

def ui_thread():
    """UI線程函數，負責顯示文件處理狀態"""
    global ui_thread_running
    
    log_message("UI線程已啟動", level='信息')
    
    try:
        # UI更新循環
        last_update_time = time.time()
        while ui_thread_running:
            try:
                current_time = time.time()
                
                # 每0.5秒更新一次UI
                if current_time - last_update_time >= 0.5:
                    display_files_status()
                    last_update_time = current_time
                
                # 使用非阻塞方式等待UI更新事件
                # 這樣即使事件沒有被設置，UI線程也能定期更新
                wait_for_ui_update()
            except Exception as e:
                log_message(f"UI線程更新時出錯: {e}", level='警告')
                # 添加短暫延遲，避免在錯誤情況下CPU使用率過高
                time.sleep(0.2)
        
        # 最後更新一次UI
        try:
            display_files_status()
        except Exception as e:
            log_message(f"最終UI更新時出錯: {e}", level='警告')
    
    finally:
        log_message("UI線程已停止", level='信息')

def start_ui_thread():
    """啟動UI線程"""
    global ui_thread_running, ui_thread_handle
    
    if ui_thread_handle is not None and ui_thread_handle.is_alive():
        log_message("UI線程已經在運行中", level='信息')
        return
    
    ui_thread_running = True
    ui_thread_handle = threading.Thread(target=ui_thread, daemon=True)
    ui_thread_handle.start()
    log_message("已啟動UI線程", level='信息')

def stop_ui_thread():
    """停止UI線程"""
    global ui_thread_running, ui_thread_handle
    
    if ui_thread_handle is None or not ui_thread_handle.is_alive():
        log_message("UI線程未運行", level='信息')
        return
    
    ui_thread_running = False
    
    # 等待UI線程結束（最多等待2秒）
    ui_thread_handle.join(2.0)
    
    if ui_thread_handle.is_alive():
        log_message("UI線程未能正常停止", level='警告')
    else:
        log_message("UI線程已停止", level='信息')
        ui_thread_handle = None