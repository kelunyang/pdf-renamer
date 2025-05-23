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

# 導入增強版模塊
from db_utils_enhanced import init_database, cleanup_database, db_connection
from log_utils_enhanced import log_message, save_log_to_csv, init_logging, cleanup_logging
from input_utils import input_helper, validate_path
from rule_utils import Rule, SimpleRule
from worker_utils import process_files_parallel, update_worker_status, ui_update_event
from pdf_utils_enhanced import extract_text_from_pdf, encrypt_pdf, split_pdf, process_pdf_files

# 全局變量
has_paddleocr = False
default_user_password = ""
default_owner_password = ""

def generate_random_password(length=8):
    """生成隨機密碼"""
    characters = string.ascii_letters + string.digits + string.punctuation
    return ''.join(random.choice(characters) for _ in range(length))

def main():
    # 初始化日誌系統
    init_logging()
    
    # 生成默認密碼
    global default_user_password, default_owner_password
    default_user_password = generate_random_password()
    default_owner_password = generate_random_password()
    
    # 檢查是否有必要的模組，若沒有則提供降級功能
    from file_utils import check_and_install_dependencies
    check_and_install_dependencies()
    
    # 導入可能在check_and_install_dependencies中安裝的模塊
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        PdfReader = PdfWriter = None
        log_message("警告: 無法導入PyPDF2，部分功能將不可用", level="警告")
    
    try:
        import fitz
    except ImportError:
        fitz = None
        log_message("警告: 無法導入PyMuPDF (fitz)，部分功能將不可用", level="警告")
    
    try:
        import pikepdf
    except ImportError:
        pikepdf = None
        log_message("警告: 無法導入pikepdf，部分功能將不可用", level="警告")
    
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
    
    log_message("歡迎使用PDF切割&重新命名小工具，本工具可以幫你把連續的PDF切割成小檔案，並根據你提供的搜尋原則重新命名／加密這些檔案")
    
    if has_paddleocr:
        log_message("【新功能】本工具現在支持對純圖片PDF進行OCR識別，可以提取圖片中的文字並保存為TXT文件，同時將OCR結果納入內容匹配流程")
    else:
        log_message("【提示】如需使用OCR功能識別純圖片PDF中的文字，請安裝PaddleOCR: pip install paddleocr paddlepaddle")
    
    # 主程序流程開始
    pdf_name = input_helper(
        "請提供要切割的檔案名稱，如果你的檔案已經切割好了，直接按Enter跳過這一步", 
        True,
        validation_func=lambda x: (not x or os.path.exists(x), "檔案不存在")
    )
    location = ""
    ori_meta = ""
    
    if pdf_name:
        location = input_helper(
            "你要把切割好的PDF檔案放在哪裡？", 
            True, 
            default="output_pdfs",
            validation_func=lambda x: (True, "")  # 目錄不需要預先存在
        )
        
        if os.path.exists(pdf_name):
            log_message("PDF檔案...已確認！")
            
            page_str = input_helper("你要幾頁切割為一個檔案？", False)
            num_page = int(page_str)
            
            if not os.path.exists(location):
                os.makedirs(location)
            
            log_message("輸出資料夾...已確認！")
            
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
                            log_message(f"警告: PDF文件 {pdf_name} 沒有有效的元數據或元數據格式不符合預期", level="警告")
                except Exception as e:
                    log_message(f"讀取元數據時出錯: {e}", level="错误")
            else:
                log_message("PyMuPDF未安裝，無法讀取PDF元數據。", level="警告")
            
            # 分割PDF - 使用可用的庫
            split_success = split_pdf(pdf_name, location, num_page, has_fitz, has_pikepdf, has_pypdf2)
            
            if not split_success:
                log_message("無法分割PDF，請確保已安裝至少一個PDF處理庫", level="错误")
                cleanup_database()
                cleanup_logging()
                sys.exit(1)
    
    # 設置搜索位置
    search_location = input_helper(
        "請提供要搜尋的資料夾位置", 
        True, 
        default=location if location else ".",
        validation_func=validate_path
    )
    
    # 設置規則
    while True:
        add_rule = input_helper(
            "是否要添加搜尋規則？(y/n)", 
            True, 
            default="y"
        ).lower() in ['y', 'yes', '']
        
        if not add_rule:
            break
        
        rule_pattern = input_helper("請輸入要搜尋的正則表達式", False)
        name = input_helper("請輸入重命名後的檔名", False)
        target_type = input_helper(
            "請選擇搜尋的目標類型 (1: 內容, 2: 檔名, 3: 元數據)", 
            True, 
            default="1"
        )
        
        target_type_map = {"1": "內容", "2": "檔名", "3": "元數據"}
        target_type = target_type_map.get(target_type, "內容")
        
        occurrence = input_helper(
            "請輸入匹配的次數 (默認為1)", 
            True, 
            default="1"
        )
        
        # 詢問是否加密
        encrypt_enable = input_helper(
            "是否要加密PDF？(y/n)", 
            True, 
            default="n"
        ).lower() in ['y', 'yes']
        
        user_pass = ""
        owner_pass = ""
        user_pass_set = True
        owner_pass_set = True
        
        if encrypt_enable:
            user_pass = input_helper(
                "請輸入開啟密碼 (留空使用隨機密碼)", 
                True, 
                is_password=True
            )
            user_pass_set = not user_pass
            
            owner_pass = input_helper(
                "請輸入編輯密碼 (留空使用隨機密碼)", 
                True, 
                is_password=True
            )
            owner_pass_set = not owner_pass
        
        # 創建規則對象
        rule_item = Rule(
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
        
        rule_items.append(rule_item)
    
    if not rule_items:
        log_message("未設置任何規則，程序將退出")
        cleanup_database()
        cleanup_logging()
        sys.exit(0)
    
    # 處理PDF文件
    log_message(f"開始處理 {search_location} 中的PDF文件...")
    
    # 獲取CPU核心數，決定最大工作線程數
    import multiprocessing
    max_workers = min(multiprocessing.cpu_count(), 4)  # 最多使用4個線程
    
    # 使用增強版pdf_utils模塊處理PDF文件
    processed_count = process_pdf_files(
        search_location, 
        rule_items, 
        has_fitz, 
        has_pypdf2, 
        has_paddleocr, 
        has_pikepdf,
        max_workers
    )
    
    log_message(f"處理完成，共處理了 {processed_count} 個文件")
    
    # 程序結束前清理資源
    cleanup_database()
    
    # 保存日誌
    cleanup_logging()
    
    print("程序執行完畢，感謝使用！")

if __name__ == "__main__":
    main()