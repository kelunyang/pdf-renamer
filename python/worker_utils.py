import os
import threading
import concurrent.futures
import queue
import time
from contextlib import contextmanager

from db_utils import db_connection, update_file_status, get_pending_files, add_files_to_database
from log_utils import log_message

# 全局變量
worker_status = {}
worker_lock = threading.Lock()
result_queue = queue.Queue()
ui_update_event = threading.Event()

@contextmanager
def worker_context(max_workers=None):
    """
    創建和管理工作線程池的上下文管理器
    
    參數:
        max_workers (int): 最大工作線程數，默認為None（使用CPU核心數）
        
    返回:
        concurrent.futures.ThreadPoolExecutor: 線程池執行器
    """
    # 如果未指定最大工作線程數，則使用CPU核心數
    if max_workers is None:
        import multiprocessing
        max_workers = multiprocessing.cpu_count()
    
    # 設置一個合理的線程數上限，避免線程過多導致程序崩潰
    # 建議線程數不超過CPU核心數的2倍，且不超過16個線程
    import multiprocessing
    cpu_count = multiprocessing.cpu_count()
    max_recommended = min(cpu_count * 2, 16)
    
    # 如果用戶設置的線程數超過建議值，則使用建議值
    if max_workers > max_recommended:
        log_message(f"警告: 設置的線程數 {max_workers} 過多，已自動調整為 {max_recommended}", level='警告')
        max_workers = max_recommended
    
    # 創建線程池
    executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
    try:
        log_message(f"已創建工作線程池，最大線程數: {max_workers}", level='信息')
        yield executor
    finally:
        # 關閉線程池
        executor.shutdown(wait=True)
        log_message("工作線程池已關閉", level='信息')

def update_worker_status(file_path, status, message=None, thread_id=None):
    """
    更新工作線程狀態
    
    參數:
        file_path (str): 文件路徑
        status (int): 狀態碼 (0: 等待, 1: 處理中, 2: 完成, 3: 失敗)
        message (str): 狀態消息
        thread_id (int): 線程ID
    """
    global worker_status, worker_lock, ui_update_event
    
    # 獲取當前時間
    current_time = time.time()
    
    # 使用鎖保護對共享字典的訪問
    with worker_lock:
        # 如果文件路徑不在狀態字典中，則添加
        if file_path not in worker_status:
            worker_status[file_path] = {
                'status': 0,  # 初始狀態為等待
                'message': '',
                'thread': None,
                'start_time': None,
                'end_time': None
            }
        
        # 更新狀態
        worker_status[file_path]['status'] = status
        
        # 如果提供了消息，則更新消息
        if message is not None:
            worker_status[file_path]['message'] = message
        
        # 如果提供了線程ID，則更新線程ID
        if thread_id is not None:
            worker_status[file_path]['thread'] = thread_id
        
        # 更新時間戳
        if status == 1 and worker_status[file_path]['start_time'] is None:  # 開始處理
            worker_status[file_path]['start_time'] = current_time
        elif status in [2, 3]:  # 完成或失敗
            worker_status[file_path]['end_time'] = current_time
    
    # 將狀態更新到數據庫
    try:
        # 使用db_utils中的函數更新文件狀態
        update_file_status(
            file_path, 
            status, 
            message, 
            thread_id, 
            worker_status[file_path]['start_time'] if status == 1 else None, 
            worker_status[file_path]['end_time'] if status in [2, 3] else None
        )
    except Exception as e:
        log_message(f"更新數據庫狀態時出錯: {e}", level='警告')
    
    # 通知UI線程更新
    ui_update_event.set()

def process_file_worker(file_path, process_func, *args, **kwargs):
    """
    工作線程處理文件的包裝函數
    
    參數:
        file_path (str): 文件路徑
        process_func (callable): 處理函數
        *args, **kwargs: 傳遞給處理函數的參數
    
    返回:
        tuple: (文件路徑, 是否成功, 結果)
    """
    # 獲取當前線程ID
    thread_id = threading.get_ident()
    
    # 更新狀態為處理中
    update_worker_status(file_path, 1, "處理中", thread_id)
    
    try:
        # 調用處理函數
        result = process_func(file_path, *args, **kwargs)
        
        # 更新狀態為完成
        update_worker_status(file_path, 2, "處理完成", thread_id)
        
        # 將結果放入結果隊列
        result_queue.put((file_path, True, result))
        
        return file_path, True, result
    except Exception as e:
        # 記錄錯誤
        error_message = f"處理文件時出錯: {e}"
        log_message(error_message, level='错误')
        
        # 更新狀態為失敗
        update_worker_status(file_path, 3, error_message, thread_id)
        
        # 將錯誤結果放入結果隊列
        result_queue.put((file_path, False, str(e)))
        
        return file_path, False, str(e)

def process_files_parallel(file_list, process_func, max_workers=None, *args, **kwargs):
    """
    並行處理多個文件
    
    參數:
        file_list (list): 文件路徑列表
        process_func (callable): 處理函數
        max_workers (int): 最大工作線程數
        *args, **kwargs: 傳遞給處理函數的參數
    
    返回:
        int: 成功處理的文件數量
    """
    if not file_list:
        log_message("沒有文件需要處理", level='信息')
        return 0
    
    # 初始化結果計數器
    success_count = 0
    total_count = len(file_list)
    
    log_message(f"開始並行處理 {total_count} 個文件...", level='信息')
    
    # 將文件添加到數據庫中
    added_count = add_files_to_database(file_list)
    log_message(f"已將 {added_count} 個文件添加到數據庫", level='信息')
    
    # 使用線程池並行處理文件
    with worker_context(max_workers) as executor:
        # 提交所有任務
        future_to_file = {}
        for file_path in file_list:
            # 更新狀態為等待
            update_worker_status(file_path, 0, "等待處理")
            
            # 提交任務
            future = executor.submit(process_file_worker, file_path, process_func, *args, **kwargs)
            future_to_file[future] = file_path
        
        # 處理結果
        for future in concurrent.futures.as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                _, success, _ = future.result()
                if success:
                    success_count += 1
            except Exception as e:
                log_message(f"獲取任務結果時出錯: {e}", level='错误')
    
    log_message(f"並行處理完成，成功: {success_count}/{total_count}", level='信息')
    return success_count

def get_worker_status():
    """
    獲取所有工作線程的狀態
    
    返回:
        dict: 工作線程狀態字典
    """
    global worker_status, worker_lock
    
    with worker_lock:
        # 返回狀態字典的副本，避免外部修改
        return worker_status.copy()

def wait_for_ui_update():
    """
    等待UI更新事件
    
    返回:
        bool: 是否有更新
    """
    global ui_update_event
    
    # 等待事件，超時時間為0.1秒
    has_update = ui_update_event.wait(0.1)
    
    # 如果有更新，則清除事件
    if has_update:
        ui_update_event.clear()
    
    return has_update