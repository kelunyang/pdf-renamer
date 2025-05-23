import sys
import os
import threading
import concurrent.futures
import gc
import time
import random
import string
import importlib.util
import queue
import csv
import shutil
import signal

# 嘗試導入questionary，如果不可用則在check_and_install_dependencies中安裝
try:
    import questionary
except ImportError:
    questionary = None

# 導入自定義模塊
from db_utils import init_database, cleanup_database, db_connection, update_file_status, get_pending_files, add_files_to_database
from log_utils import log_message, save_log_to_csv, log_entries
from input_utils import input_helper, validate_path
from rule_utils import Rule, SimpleRule
from file_utils import file_renamer, check_and_install_dependencies
from pdf_utils import extract_text_from_pdf, encrypt_pdf, split_pdf, process_pdf_files

# 全局變量
has_paddleocr = False
dict_lock = threading.Lock()
pdf_status_dict = {}
# 移除ui_update_event的初始化，不再使用事件通知機制
default_user_password = ""
default_owner_password = ""
worker_status = {}
worker_lock = threading.Lock()
result_queue = queue.Queue()
# OCR相關全局設置
remove_whitespace = True  # 默認去除OCR結果中的空白

# 設置PaddleOCR的日誌重定向
try:
    from paddle_utils import setup_paddle_logging
    setup_paddle_logging()
    log_message("已設置PaddleOCR的日誌重定向", level='信息')
except ImportError:
    log_message("paddle_utils模塊不可用，PaddleOCR的警告訊息將顯示在控制台", level='警告')

def generate_random_password(length=8):
    """生成隨機密碼"""
    characters = string.ascii_letters + string.digits + string.punctuation
    return ''.join(random.choice(characters) for _ in range(length))

def get_terminal_size():
    """獲取終端大小"""
    try:
        from shutil import get_terminal_size as get_size
        return get_size()
    except (ImportError, AttributeError):
        # 如果無法獲取終端大小，返回默認值
        from collections import namedtuple
        Size = namedtuple('Size', ['columns', 'lines'])
        return Size(80, 24)

def clear_screen():
    """清除屏幕"""
    os.system('cls' if os.name == 'nt' else 'clear')

def truncate_filename(filename):
    """縮短檔名顯示：首尾各3個字符，中間用...代替
    
    參數:
        filename (str): 完整檔名
    
    返回:
        str: 縮短後的檔名
    """
    # 分離檔名和副檔名
    name, ext = os.path.splitext(filename)
    
    # 如果檔名長度小於等於6，直接返回原檔名
    if len(name) <= 6:
        return filename
    
    # 取首尾各3個字符，中間用...代替
    return f"{name[:3]}...{name[-3:]}{ext}"

def display_files_status():
    """顯示文件處理狀態，只顯示錯誤訊息而非所有更名操作"""
    clear_screen()
    print("檔案更名狀態:\n")
    
    # 獲取終端大小
    terminal_size = get_terminal_size()
    terminal_width = terminal_size.columns
    
    # 計算每行可以顯示的文件數量
    files_per_line = max(1, terminal_width // 30)  # 假設每個文件名平均30個字符
    
    try:
        # 按狀態分類文件
        processing_files = []  # 處理中（黃色）
        queued_files = []      # 排隊中（白色）
        failed_files = []      # 失敗（紅色）
        
        # 獲取錯誤日誌（最近的5條）
        error_logs = []
        for entry in log_entries:
            if entry.get('level', '') == '错误':
                error_logs.append(entry)
        # 按時間倒序排序並只保留最近的5條
        error_logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        error_logs = error_logs[:5]
        
        # 從數據庫獲取文件狀態
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT path, status, message, thread FROM files')
            rows = cursor.fetchall()
            
            # 計算統計信息
            pending = 0
            queued = 0
            processing = 0
            failed = 0
            completed = 0
            
            for row in rows:
                path, status, message, thread_id = row
                file_name = os.path.basename(path)
                
                if status == 0:  # 待處理
                    # 檢查是否已排入隊列
                    if thread_id is not None:
                        queued += 1
                        # 縮短檔名顯示：首尾各3個字符，中間用...代替
                        short_name = truncate_filename(file_name)
                        queued_files.append(f"{short_name}")  # 白色
                    else:
                        pending += 1
                elif status == 1:  # 處理中
                    processing += 1
                    # 縮短檔名顯示：首尾各3個字符，中間用...代替
                    short_name = truncate_filename(file_name)
                    processing_files.append(f"\033[93m{short_name}\033[0m")  # 黃色
                elif status == 2:  # 已處理成功
                    completed += 1
                elif status == 3:  # 處理失敗
                    failed += 1
                    # 縮短檔名顯示：首尾各3個字符，中間用...代替
                    short_name = truncate_filename(file_name)
                    failed_files.append(f"\033[91m{short_name} ({message})\033[0m")  # 紅色
        
        total = pending + queued + processing + completed + failed
        
        # 顯示處理中的文件 (優先顯示)
        if processing_files:
            print("處理中:")
            for i in range(0, len(processing_files), files_per_line):
                end_idx = min(i + files_per_line, len(processing_files))
                print(" ".join(processing_files[i:end_idx]))
            print()
        
        # 顯示排隊中的文件
        if queued_files:
            print("排隊中:")
            for i in range(0, len(queued_files), files_per_line):
                end_idx = min(i + files_per_line, len(queued_files))
                print(" ".join(queued_files[i:end_idx]))
            print()
        
        # 顯示失敗的文件
        if failed_files:
            print("處理失敗:")
            for i in range(0, len(failed_files), files_per_line):
                end_idx = min(i + files_per_line, len(failed_files))
                print(" ".join(failed_files[i:end_idx]))
            print()
        
        # 使用tqdm顯示進度條
        if total > 0:
            try:
                # 導入tqdm
                from tqdm import tqdm
                # 創建tqdm進度條
                progress = completed / total
                # 計算已經過的時間（秒）
                elapsed_time = 1.0  # 使用一個默認值，避免使用None
                # 使用tqdm.format_meter生成進度條字符串
                bar = tqdm.format_meter(
                    n=completed,
                    total=total,
                    elapsed=elapsed_time,
                    ncols=terminal_width - 20,
                    prefix="進度:"
                )
                print(f"{bar}\n")
            except ImportError:
                # 如果tqdm不可用，使用原始進度條
                progress = completed / total
                bar_width = terminal_width - 20
                filled_width = int(bar_width * progress)
                bar = f"[{'#' * filled_width}{'-' * (bar_width - filled_width)}] {progress:.1%}"
                print(f"進度: {bar}\n")
        
        # 顯示錯誤日誌（紅色，按時間倒序）
        print("最近錯誤:")
        if error_logs:
            for entry in error_logs:
                timestamp = entry.get('timestamp', '')
                message = entry.get('message', '')
                print(f"\033[91m[{timestamp}] {message}\033[0m")
        else:
            print("無錯誤日誌")
        print()
        
        # 顯示統計信息
        print(f"統計: 總計 {total} 個檔案, 已完成 {completed}, 處理中 {processing}, 已排入隊列 {queued}, 等待中 {pending}, 失敗 {failed}")
    
    except Exception as e:
        log_message(f"顯示檔案狀態時出錯: {e}", level='警告')

def ui_thread_function():
    """UI線程函數，用於顯示文件處理狀態"""
    while True:
        # 不再等待更新事件，改為直接顯示文件狀態
        display_files_status()
        
        # 檢查是否所有文件都已處理完成
        with db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM files')
            total = cursor.fetchone()[0]
            
            if total > 0:
                cursor.execute('SELECT COUNT(*) FROM files WHERE status = 2')
                completed = cursor.fetchone()[0]
                
                # 如果所有文件都已處理完成（進度100%），則停止更新UI
                if completed == total:
                    # 最後更新一次UI
                    display_files_status()
                    log_message("所有文件處理完成，UI更新停止", level='信息')
                    break
        
        # 休眠1秒，定期查詢數據庫更新UI
        time.sleep(1.0)

def update_worker_status(file_path, status, message=None, thread_id=None):
    """更新工作線程狀態"""
    global worker_status, worker_lock
    
    # 獲取當前時間
    current_time = time.time()
    
    # 使用鎖保護共享數據的訪問
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
    
    # 不再通知UI線程更新，UI更新將完全依賴於定時查詢數據庫

def process_file_worker(file_path, process_func, *args, **kwargs):
    """工作線程處理文件的包裝函數"""
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
    """並行處理多個文件"""
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
    
    # 使用worker_context來創建和管理線程池，確保線程數量受到控制
    from worker_utils import worker_context
    with worker_context(max_workers=max_workers) as executor:
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

def import_rules_from_csv(csv_path):
    """從CSV文件導入規則
    CSV格式: 關鍵字,目標檔名,原則,重複次數,開啟密碼,編輯密碼
    """
    rules = []
    try:
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 6:
                    # 解析CSV行
                    rule_pattern = row[0].strip()
                    name = row[1].strip()
                    target_type = row[2].strip()
                    occurrence = row[3].strip() if len(row) > 3 and row[3].strip() else "1"
                    user_pass = row[4].strip() if len(row) > 4 else ""
                    owner_pass = row[5].strip() if len(row) > 5 else ""
                    
                    # 處理b''格式的字節字符串
                    def convert_byte_str(s):
                        if s.startswith("b'") and s.endswith("'"):
                            return s[2:-1].encode('utf-8').decode('unicode_escape')
                        return s
                    
                    user_pass = convert_byte_str(user_pass)
                    owner_pass = convert_byte_str(owner_pass)
                    
                    # 修正加密判斷邏輯
                    encrypt_enable = bool(user_pass.strip() or owner_pass.strip())
                    user_pass_set = not user_pass.strip()
                    owner_pass_set = not owner_pass.strip()
                    
                    # 創建規則對象
                    rule = Rule(
                        rule_pattern, 
                        name, 
                        target_type, 
                        occurrence, 
                        user_pass, 
                        owner_pass, 
                        user_pass_set, 
                        owner_pass_set, 
                        encrypt_enable
                    )
                    rules.append(rule)
        print(f"成功從CSV導入了 {len(rules)} 條規則")
    except Exception as e:
        print(f"導入CSV規則時出錯: {e}")
    return rules

def signal_handler(sig, frame):
    """處理Ctrl+C信號"""
    print("\n\n收到中斷信號，正在結束程序...")
    log_message("收到中斷信號，正在結束程序...", level='警告')
    
    # 清理資源
    try:
        cleanup_database()
        log_message("已清理數據庫", level='信息')
    except Exception as e:
        log_message(f"清理數據庫時出錯: {e}", level='警告')
    
    # 保存日誌
    try:
        save_log_to_csv()
        log_message("已保存日誌", level='信息')
    except Exception as e:
        log_message(f"保存日誌時出錯: {e}", level='警告')
    
    print("程序已安全結束，感謝使用！")
    sys.exit(0)

def preview_pdf_content(pdf_path, num_pages=None, use_ocr=False, remove_whitespace=False, save_txt=False):
    """預覽PDF文件內容
    
    參數:
        pdf_path (str): PDF文件路徑
        num_pages (int): 要預覽的頁數，None表示全部
        use_ocr (bool): 是否使用OCR識別文字
        remove_whitespace (bool): 是否去除中文空白
        save_txt (bool): 是否將OCR結果保存為txt文件
    """
    try:
        # 如果啟用OCR且有PaddleOCR，則使用OCR預覽
        global has_paddleocr
        if use_ocr and has_paddleocr:
            try:
                from pdf_utils import extract_text_from_pdf
                import fitz
                
                # 檢查是否有PyMuPDF
                if not importlib.util.find_spec("fitz"):
                    print("無法使用OCR預覽：缺少必要的庫（PyMuPDF）")
                    return False
                
                log_message(f"正在處理PDF文件進行預覽: {os.path.basename(pdf_path)}", level='信息')
                print(f"\n使用OCR預覽文件: {os.path.basename(pdf_path)}")
                print("正在處理中，請稍候...\n")
                
                # 使用OCR提取文本
                ocr_text = extract_text_from_pdf(pdf_path, has_fitz=True, has_pypdf2=False, has_paddleocr=True, 
                                               force_ocr=True, remove_whitespace=remove_whitespace, save_txt=save_txt,
                                               preview_mode=True)
                
                # 按頁分割文本
                pages = ocr_text.split("===== 第")
                if len(pages) > 1:
                    pages = pages[1:]  # 跳過第一個空元素
                    
                    # 限制預覽頁數
                    pages_to_preview = len(pages) if num_pages is None else min(num_pages, len(pages))
                    
                    print(f"總頁數: {len(pages)}，預覽頁數: {pages_to_preview}\n")
                    print("-" * 80)
                    
                    for i in range(pages_to_preview):
                        page_text = pages[i]
                        # 提取頁碼
                        page_num = page_text.split(" ")[0].strip()
                        # 提取內容
                        content = page_text.split("=====")[1] if "=====" in page_text else page_text
                        
                        print(f"第 {page_num} 頁:")
                        # 只顯示前500個字符
                        if len(content) > 500:
                            content = content[:500] + "...（內容已截斷）"
                        print(content)
                        print("-" * 80)
                else:
                    # 如果無法按頁分割，則顯示整個文本
                    print("無法按頁分割OCR結果，顯示完整文本:")
                    print("-" * 80)
                    # 只顯示前1000個字符
                    if len(ocr_text) > 1000:
                        ocr_text = ocr_text[:1000] + "...（內容已截斷）"
                    print(ocr_text)
                    print("-" * 80)
                
                return True
            except Exception as e:
                print(f"使用OCR預覽PDF內容時出錯: {e}")
                print("嘗試使用標準方法預覽...")
        
        # 嘗試使用PyMuPDF (fitz)
        try:
            import fitz
            with fitz.open(pdf_path) as doc:
                total_pages = len(doc)
                pages_to_preview = total_pages if num_pages is None else min(num_pages, total_pages)
                
                print(f"\n預覽文件: {os.path.basename(pdf_path)}")
                print(f"總頁數: {total_pages}，預覽頁數: {pages_to_preview}\n")
                print("-" * 80)
                
                for i in range(pages_to_preview):
                    print(f"第 {i+1} 頁:")
                    text = doc[i].get_text()
                    # 根據設置決定是否去除空白
                    if remove_whitespace:
                        text = text.replace(" ", "")
                    # 只顯示前500個字符
                    if len(text) > 500:
                        text = text[:500] + "...（內容已截斷）"
                    print(text)
                    print("-" * 80)
                
                return True
        except ImportError:
            pass
        
        # 嘗試使用PyPDF2
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            pages_to_preview = total_pages if num_pages is None else min(num_pages, total_pages)
            
            print(f"\n預覽文件: {os.path.basename(pdf_path)}")
            print(f"總頁數: {total_pages}，預覽頁數: {pages_to_preview}\n")
            print("-" * 80)
            
            for i in range(pages_to_preview):
                print(f"第 {i+1} 頁:")
                text = reader.pages[i].extract_text()
                # 根據設置決定是否去除空白
                if remove_whitespace:
                    text = text.replace(" ", "")
                # 只顯示前500個字符
                if len(text) > 500:
                    text = text[:500] + "...（內容已截斷）"
                print(text)
                print("-" * 80)
            
            return True
        except ImportError:
            print("無法預覽PDF內容：缺少必要的庫（PyMuPDF或PyPDF2）")
            return False
    except Exception as e:
        print(f"預覽PDF內容時出錯: {e}")
        return False

def main():
    # 註冊信號處理器，捕捉Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)
    
    # 同時在file_utils模塊中註冊信號處理器
    from file_utils import file_utils_signal_handler
    # 在Windows系統中，信號處理只在主線程中有效
    # 因此我們需要在多個地方註冊信號處理器
    signal.signal(signal.SIGINT, file_utils_signal_handler)
    
    # 生成默認密碼
    global default_user_password, default_owner_password
    default_user_password = generate_random_password()
    default_owner_password = generate_random_password()
    
    # 檢查是否有必要的模組，若沒有則提供降級功能
    check_and_install_dependencies()
    
    # 重新嘗試導入questionary（可能在check_and_install_dependencies中安裝）
    global questionary
    if questionary is None:
        try:
            import questionary
        except ImportError:
            print("警告: 無法導入questionary，部分交互功能將降級使用標準輸入")
            # 創建一個簡單的替代品
            class QuestionaryMock:
                @staticmethod
                def confirm(message, default=True):
                    class ConfirmMock:
                        @staticmethod
                        def ask():
                            print(f"{message} (y/n) [{'y' if default else 'n'}]: ")
                            response = input().strip().lower()
                            if not response:
                                return default
                            return response in ['y', 'yes']
                    return ConfirmMock()
            questionary = QuestionaryMock()
    
    # 導入可能在check_and_install_dependencies中安裝的模塊
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        PdfReader = PdfWriter = None
        print("警告: 無法導入PyPDF2，部分功能將不可用")
    
    try:
        import fitz
    except ImportError:
        fitz = None
        print("警告: 無法導入PyMuPDF (fitz)，部分功能將不可用")
    
    try:
        import pikepdf
    except ImportError:
        pikepdf = None
        print("警告: 無法導入pikepdf，部分功能將不可用")
    
    try:
        from tqdm import tqdm
    except ImportError:
        # 如果tqdm不可用，創建一個簡單的替代品
        def tqdm(iterable, **kwargs):
            return iterable
    
    has_pikepdf = pikepdf is not None
    has_fitz = fitz is not None
    has_pypdf2 = PdfReader is not None
    global has_paddleocr
    has_paddleocr = importlib.util.find_spec("paddleocr") is not None and importlib.util.find_spec("paddle") is not None
    
    # 初始化數據庫
    init_database()
    
    rule_items = []
    
    print("歡迎使用PDF切割&重新命名小工具，本工具可以幫你把連續的PDF切割成小檔案，並根據你提供的搜尋原則重新命名／加密這些檔案")
    
    if has_paddleocr:
        print("【新功能】本工具現在支持對純圖片PDF進行OCR識別，可以提取圖片中的文字並保存為TXT文件，同時將OCR結果納入內容匹配流程")
    else:
        print("【提示】如需使用OCR功能識別純圖片PDF中的文字，請安裝PaddleOCR: pip install paddleocr paddlepaddle")
    
    # 選擇操作模式
    operation_mode = questionary.select(
        "請選擇操作模式：",
        choices=[
            "切割PDF並重命名",
            "從CSV導入規則並重命名"
        ]
    ).ask()
    
    # 設置全局密碼
    global_user_pass = input_helper(
        "請設置全局PDF開啟密碼（留空使用隨機密碼[ " + default_user_password + " ]）", 
        True, 
        is_password=True
    )
    if global_user_pass:
        default_user_password = global_user_pass
    
    global_owner_pass = input_helper(
        "請設置全局PDF編輯密碼（留空使用隨機密碼[ " + default_owner_password + " ]）", 
        True, 
        is_password=True
    )
    if global_owner_pass:
        default_owner_password = global_owner_pass
    
    # 主程序流程開始
    pdf_name = ""
    location = ""
    ori_meta = ""
    
    # 根據操作模式處理
    if operation_mode == "切割PDF並重命名":
        pdf_name = input_helper(
            "請提供要切割的檔案名稱", 
            False,
            validation_func=lambda x: (os.path.exists(x), "檔案不存在")
        )
    
    if pdf_name:
        location = input_helper(
            "你要把切割好的PDF檔案放在哪裡？", 
            False, 
            validation_func=lambda x: (os.path.exists(x), "目錄不存在")
        )
        
        if os.path.exists(pdf_name):
            print("PDF檔案...已確認！")
            
            page_str = input_helper("你要幾頁切割為一個檔案？", False)
            num_page = int(page_str)
            
            if not os.path.exists(location):
                os.makedirs(location)
            
            print("輸出資料夾...已確認！")
            
            # 讀取PDF元數據
            if has_fitz:
                try:
                    with fitz.open(pdf_name) as doc:
                        # 添加防呆機制，確保metadata是字典類型且不為None
                        if doc.metadata and isinstance(doc.metadata, dict):
                            for key, value in doc.metadata.items():
                                if value:
                                    ori_meta += str(value)
                        else:
                            print(f"警告: PDF文件 {pdf_name} 沒有有效的元數據或元數據格式不符合預期")
                except Exception as e:
                    print(f"讀取元數據時出錯: {e}")
            else:
                print("PyMuPDF未安裝，無法讀取PDF元數據。")
            
            # 分割PDF - 使用可用的庫
            split_success = split_pdf(pdf_name, location, num_page, has_fitz, has_pikepdf, has_pypdf2)
            
            if not split_success:
                print("無法分割PDF，請確保已安裝至少一個PDF處理庫")
                sys.exit(1)
    
    # 設置搜索位置
    search_location = input_helper(
        "請提供要搜尋的資料夾位置", 
        True, 
        default=location if location else ".",
        validation_func=validate_path
    )
    
    # 獲取PDF文件列表
    pdf_files = []
    for root, _, files in os.walk(search_location):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                pdf_files.append(file_path)
    
    # 獲取當前目錄下的PDF文件列表（用於預覽）
    local_pdf_files = [f for f in os.listdir(search_location) if f.lower().endswith('.pdf')]
    
    # 詢問是否啟用OCR功能
    use_ocr = False
    save_ocr_txt = False
    
    if has_paddleocr:
        if questionary:
            use_ocr = questionary.select(
                "是否對所有PDF使用OCR功能？(即使PDF包含文字也會使用OCR提取內容，適用於掃描文件)",
                choices=["是", "否"],
                default="否"
            ).ask() == "是"
        else:
            use_ocr = input_helper(
                "是否對所有PDF使用OCR功能？(y/n)\n(即使PDF包含文字也會使用OCR提取內容，適用於掃描文件)",
                True, 
                default="n"
            ).lower() in ['y', 'yes']
        
        if use_ocr:
            # 只寫入日誌，不顯示訊息
            log_message("已啟用OCR功能，將對所有PDF文件進行OCR處理", level='信息')
            
            # 詢問是否去除OCR結果中的空白
            if questionary:
                remove_whitespace = questionary.select(
                "是否去除OCR結果中的空白？(適合中文文檔)",
                    choices=["是", "否"],
                    default="是"
                ).ask() == "是"
            else:
                remove_whitespace = input_helper(
                    "是否去除OCR結果中的空白？(y/n)\n(適合中文文檔)",
                    True,
                    default="y"
                ).lower() in ['y', 'yes']
            
            # 詢問是否將OCR結果輸出為txt檔案
            if questionary:
                save_ocr_txt = questionary.select(
                "是否將OCR結果輸出為單獨的txt檔案？",
                    choices=["是", "否"],
                    default="是"
                ).ask() == "是"
            else:
                save_ocr_txt = input_helper(
                    "是否將OCR結果輸出為單獨的txt檔案？(y/n)",
                    True,
                    default="y"
                ).lower() in ['y', 'yes']
            
        else:
            log_message("OCR功能未啟用，僅在檢測到純圖片PDF時才會使用OCR", level='信息')
    else:
        log_message("未安裝PaddleOCR，無法使用OCR功能", level='警告')
    
    # 從CSV導入規則
    if operation_mode == "從CSV導入規則並重命名" or pdf_name:
        csv_path = input_helper(
            "請提供規則CSV文件的路徑",
            False,
            validation_func=lambda x: (os.path.exists(x) and x.lower().endswith('.csv'), "檔案不存在或不是CSV文件")
        )
        rule_items = import_rules_from_csv(csv_path)
    
    # 預覽第一個PDF文件
    if local_pdf_files:
        if questionary:
            preview_choice = questionary.select(
            f"是否要預覽第一個PDF文件 ({local_pdf_files[0]}) 的內容？",
                choices=["是", "否"],
                default="是"
            ).ask() == "是"
        else:
            preview_choice = input_helper(
                f"是否要預覽第一個PDF文件 ({local_pdf_files[0]}) 的內容？(y/n)",
                True,
                default="y",
                is_confirm=True
            )
        
        if preview_choice:
            # 詢問是否使用OCR預覽
            use_ocr_preview = False
            if has_paddleocr:
                if questionary:
                    use_ocr_preview = questionary.select(
                    "是否使用OCR預覽PDF內容？(適用於掃描文件)",
                        choices=["是", "否"],
                        default="否"
                    ).ask() == "是"
                else:
                    use_ocr_preview = input_helper(
                        "是否使用OCR預覽PDF內容？(適用於掃描文件)(y/n)",
                        True,
                        default="n",
                        is_confirm=True
                    )
                
                # 如果使用OCR，直接使用全局設置的空白處理選項
                if use_ocr_preview:
                    # 使用全局設置，不再詢問用戶
                    remove_whitespace_preview = remove_whitespace
                    # 只寫入日誌，不顯示訊息
                    log_message(f"OCR預覽使用全局空白處理設置: {'去除空白' if remove_whitespace_preview else '保留空白'}", level='信息')
            
            preview_pages = input_helper(
                "請輸入要預覽的頁數（按Enter預覽全部）",
                True,
                default=""
            )
            
            num_pages = None
            if preview_pages:
                try:
                    num_pages = int(preview_pages)
                except ValueError:
                    print("無效的頁數，將預覽全部頁面")
            
            # 使用与OCR设置相同的参数进行预览
            if use_ocr and use_ocr_preview:
                # 如果全局OCR已启用且预览也使用OCR，使用相同的设置
                preview_pdf_content(os.path.join(search_location, local_pdf_files[0]), num_pages, use_ocr_preview, remove_whitespace, save_ocr_txt)
            else:
                # 否则使用预览特定的设置
                preview_pdf_content(os.path.join(search_location, local_pdf_files[0]), num_pages, use_ocr_preview, remove_whitespace_preview, False)
    
    if not rule_items:
        print("未設置任何規則，程序將退出")
        sys.exit(0)
    
    # 處理PDF文件
    print(f"開始處理 {search_location} 中的PDF文件...")
    
    # 獲取PDF文件列表
    pdf_files = []
    for root, _, files in os.walk(search_location):
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                pdf_files.append(file_path)
    
    if not pdf_files:
        print("未找到PDF文件")
        return
    
    print(f"找到 {len(pdf_files)} 個PDF文件")
    
    # 顯示OCR設置信息
    if use_ocr:
        print(f"OCR功能: 已啟用")
        print(f"去除空白: {'已啟用' if remove_whitespace else '未啟用'}")
        print(f"保存OCR文本: {'已啟用' if save_ocr_txt else '未啟用'}")
    else:
        print("OCR功能: 未啟用")
    
    # 獲取CPU核心數，決定最大工作線程數
    import multiprocessing
    cpu_count = multiprocessing.cpu_count()
    
    # 詢問用戶希望使用的線程數
    thread_count_str = input_helper(
        f"請輸入最大工作線程數 (1-{cpu_count}，按Enter使用默認值4): ",
        True,
        default="4",
        validation_func=lambda x: (x.isdigit() and 1 <= int(x) <= cpu_count, f"請輸入1到{cpu_count}之間的數字")
    )
    
    # 解析用戶輸入的線程數
    max_workers = min(int(thread_count_str), cpu_count)  # 確保不超過CPU核心數
    print(f"將使用 {max_workers} 個執行緒處理PDF檔案")
    
    # 詢問操作類型（重命名或複製）
    if questionary:
        operation_type = questionary.select(
            "請選擇操作類型：",
            choices=[
                "重命名文件",
                "複製文件（保留原文件）"
            ],
            default="重命名文件"
        ).ask()
    else:
        print("請選擇操作類型：")
        print("1. 重命名文件")
        print("2. 複製文件（保留原文件）")
        choice = input_helper("請輸入選項編號(1/2): ", True, default="1")
        operation_type = "複製文件（保留原文件）" if choice == "2" else "重命名文件"
    
    is_copy_mode = operation_type == "複製文件（保留原文件）"
    
    # 確認是否開始處理
    if questionary:
        start_processing = questionary.select(
            f"是否開始執行{'複製' if is_copy_mode else '更名'}程序？",
                choices=["是", "否"],
                default="是"
            ).ask() == "是"
    else:
        start_processing = input_helper(
            f"是否開始執行{'複製' if is_copy_mode else '更名'}程序？(y/n)", 
            True, 
            default="y"
        ).lower() in ['y', 'yes']
    
    if not start_processing:
        print("操作已取消")
        return
    
    # 啟動UI線程
    ui_thread = threading.Thread(target=ui_thread_function, daemon=True)
    ui_thread.start()
    print("已啟動UI線程，將顯示處理進度...")
    
    # 給UI線程一點時間來初始化顯示
    time.sleep(0.5)
    
    # 使用之前已經設置的OCR選項
    if use_ocr:
        log_message("使用已設置的OCR選項處理所有PDF文件", level='信息')
    else:
        log_message("OCR功能未啟用或不可用，僅在檢測到純圖片PDF時才會使用OCR", level='信息')
    
    # 使用pdf_utils模塊處理PDF文件，並使用並行處理
    processed_count = process_pdf_files(
        pdf_files, 
        rule_items, 
        search_location, 
        ori_meta, 
        has_fitz, 
        has_pypdf2, 
        has_paddleocr, 
        has_pikepdf, 
        max_workers,
        is_copy_mode,
        use_ocr,  # 傳遞OCR選項
        remove_whitespace,  # 傳遞去除空白選項
        save_ocr_txt,  # 傳遞保存OCR文本選項
        default_user_password,  # 傳遞默認用戶密碼
        default_owner_password  # 傳遞默認所有者密碼
    )
    
    # 如果啟用了OCR和保存TXT，提示用戶
    if use_ocr and save_ocr_txt:
        print(f"\nOCR處理完成，文本文件已保存在PDF文件所在目錄\n")
    
    # 最後更新一次UI
    # 處理完成，給UI線程一點時間來最後更新一次
    time.sleep(1)
    
    print(f"處理完成，共處理了 {processed_count} 個文件")
    
    # 程序結束前清理資源
    cleanup_database()
    
    # 保存日誌
    try:
        save_log_to_csv()
    except Exception as e:
        log_message(f"保存日誌時出錯: {e}", level='警告')
    
    # 詢問是否查看處理日誌
    if questionary.select(
        "是否要查看處理日誌？",
        choices=["是", "否"],
        default="否"
    ).ask() == "是":
        print("\n處理日誌:")
        for entry in log_entries:
            print(f"[{entry.get('timestamp', '')}] [{entry.get('level', '未知')}] {entry.get('message', '')}")
    
    print("程序執行完畢，感謝使用！")

def cleanup_logging():
    """清理日誌系統並保存日誌"""
    try:
        save_log_to_csv()
    except Exception as e:
        log_message(f"保存日誌時出錯: {e}", level='警告')

if __name__ == "__main__":
    main()