import time
import os
import threading

# 初始化全局日志列表
log_entries = []
log_lock = threading.Lock()

def log_message(message, level='信息', thread_name=None):
    """
    记录消息到全局日志列表（线程安全）
    
    参数:
        message (str): 要记录的消息
        level (str): 日志级别，默认为'信息'
        thread_name (str): 线程名称，默认为None（自动获取）
    """
    global log_entries, log_lock
    
    # 如果未提供线程名称，则获取当前线程名称
    if thread_name is None:
        thread_name = threading.current_thread().name
    
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"{timestamp} [{level}] [{thread_name}] {message}"
    
    # 使用锁保护对共享列表的访问
    with log_lock:
        log_entries.append({
            'timestamp': timestamp,
            'level': level,
            'thread': thread_name,
            'message': message
        })
    
    #print(log_entry)
    
    # 通知UI线程更新
    # 移除对ui_update_event的导入和使用
    try:
        # 尝试通知UI更新，但不再依赖ui_update_event
        # 不再从main导入ui_update_event
        # 不再调用ui_update_event.set()
        pass
    except Exception as e:
        pass  # 如果通知UI更新失败，则忽略

def init_logging():
    """
    初始化日誌系統
    """
    global log_entries, log_lock
    
    # 清空日誌列表
    with log_lock:
        log_entries = []
    
    log_message("日誌系統已初始化")

def cleanup_logging():
    """
    清理日誌系統並保存日誌
    """
    save_log_to_csv()
    
    # 清空日誌列表釋放內存
    global log_entries, log_lock
    with log_lock:
        log_entries = []

def log_exception(exception, message="發生異常"):
    """
    記錄異常信息
    
    參數:
        exception (Exception): 異常對象
        message (str): 附加消息
    """
    log_message(f"{message}: {str(exception)}", level="错误")

def save_log_to_csv():
    """
    將日誌保存到CSV文件
    """
    import csv
    global log_entries, log_lock
    
    # 使用鎖保護對共享列表的訪問
    with log_lock:
        if not log_entries:
            print("沒有日誌記錄需要保存")
            return
        
        # 創建日誌副本以避免在寫入過程中被修改
        entries_copy = log_entries.copy()
    
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    log_filename = f"pdfRenamer_log_{timestamp}.csv"
    
    try:
        # 使用utf-8-sig編碼處理BOM問題
        with open(log_filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['時間戳', '級別', '線程', '消息'])
            
            for entry in entries_copy:
                csv_writer.writerow([
                    entry.get('timestamp', ''),
                    entry.get('level', '未知'),
                    entry.get('thread', '未知'),
                    entry.get('message', '')
                ])
            
        print(f"日誌已保存到 {log_filename}")
    except Exception as e:
        print(f"保存日誌時出錯: {e}")