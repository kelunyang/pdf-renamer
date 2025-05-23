import os
import gc
import time
import sys
import re
import importlib.util
from log_utils import log_message

def extract_text_from_pdf(pdf_path, has_fitz=False, has_pypdf2=False, has_paddleocr=False, force_ocr=False, remove_whitespace=False, save_txt=False, output_txt_path=None, ocr_instance=None, preview_mode=False):
    """從PDF文件中提取文本
    
    參數:
        pdf_path (str): PDF文件路徑
        has_fitz (bool): 是否有PyMuPDF
        has_pypdf2 (bool): 是否有PyPDF2
        has_paddleocr (bool): 是否有PaddleOCR
        force_ocr (bool): 是否強制使用OCR，無論是否有文本
        remove_whitespace (bool): 是否去除OCR結果中的空白
        save_txt (bool): 是否將OCR結果保存為txt文件
        output_txt_path (str): OCR結果保存路徑，如果為None則使用原始檔名
        ocr_instance (PaddleOCR): 可選的PaddleOCR實例，如果提供則使用此實例
        preview_mode (bool): 是否為預覽模式，預覽模式下不保存TXT文件
        
    返回:
        str: 提取的文本
    """
    text = ""
    
    # 如果強制使用OCR且有PaddleOCR，則直接使用OCR
    if force_ocr and has_paddleocr and has_fitz:
        return extract_text_with_paddleocr(pdf_path, remove_whitespace, save_txt, output_txt_path, ocr_instance, preview_mode)
    
    # 嘗試使用PyMuPDF提取文本
    if has_fitz:
        try:
            import fitz
            with fitz.open(pdf_path) as doc:
                for page in doc:
                    text += page.get_text()
            # 如果提取到文本且不強制使用OCR，則返回
            if text and not force_ocr:
                return text
        except Exception as e:
            print(f"使用PyMuPDF提取文本時出錯: {e}")
    
    # 嘗試使用PyPDF2提取文本
    if has_pypdf2 and not text and not force_ocr:
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text
            # 如果提取到文本且不強制使用OCR，則返回
            if text and not force_ocr:
                return text
        except Exception as e:
            print(f"使用PyPDF2提取文本時出錯: {e}")
    
    # 如果沒有提取到文本或強制使用OCR，且有PaddleOCR，則使用OCR
    if has_paddleocr and (not text or force_ocr):
        ocr_text = extract_text_with_paddleocr(pdf_path, remove_whitespace, save_txt, output_txt_path, ocr_instance, preview_mode)
        if ocr_text:
            return ocr_text
    
    return text

def extract_text_with_paddleocr(pdf_path, remove_whitespace=False, save_txt=False, output_txt_path=None, ocr_instance=None, preview_mode=False):
    """使用PaddleOCR從PDF提取文本
    
    參數:
        pdf_path (str): PDF文件路徑
        remove_whitespace (bool): 是否去除OCR結果中的空白
        save_txt (bool): 是否將OCR結果保存為txt文件
        output_txt_path (str): OCR結果保存路徑，如果為None則使用原始檔名
        ocr_instance (PaddleOCR): 可選的PaddleOCR實例，如果提供則使用此實例
        preview_mode (bool): 是否為預覽模式，預覽模式下不保存TXT文件
        
    返回:
        str: 提取的文本
    """
    try:
        import fitz
        import os
        import tempfile
        import numpy as np
        from PIL import Image
        
        # 使用提供的OCR實例或通過paddle_utils創建新的實例
        if ocr_instance:
            ocr = ocr_instance
        else:
            # 導入自定義的paddle_utils模塊來初始化PaddleOCR
            try:
                from paddle_utils import init_paddleocr
                ocr = init_paddleocr(use_angle_cls=True, lang="ch")
                if ocr is None:
                    # 如果初始化失敗，嘗試直接導入PaddleOCR
                    from paddleocr import PaddleOCR
                    ocr = PaddleOCR(use_angle_cls=True, lang="ch")
            except ImportError:
                # 如果paddle_utils不可用，直接導入PaddleOCR
                from paddleocr import PaddleOCR
                ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        ocr_text = ""
        
        # 創建臨時目錄用於存儲圖片
        temp_dir = tempfile.mkdtemp()
        
        try:
            # 打開PDF文件
            with fitz.open(pdf_path) as doc:
                print(f"使用PaddleOCR處理PDF: {os.path.basename(pdf_path)}，共{len(doc)}頁")
                
                # 處理每一頁
                for page_num, page in enumerate(doc):
                    # 計算適當的縮放比例以獲得300dpi的解析度
                    # 標準PDF點數為72dpi，所以縮放比例為300/72 = 4.167
                    # 但由於PyMuPDF的Matrix縮放可能會導致圖像過大，我們使用較低的值
                    scale_factor = 300 / 72
                    
                    # 將頁面渲染為圖片，使用300dpi的解析度並轉為灰階
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor), colorspace="gray")
                    img_path = os.path.join(temp_dir, f"temp_page_{page_num}.png")
                    pix.save(img_path)
                    
                    print(f"正在OCR處理第{page_num+1}頁...")
                    
                    # 使用OCR識別圖片中的文字
                    try:
                        result = ocr.ocr(img_path, cls=True)
                    except Exception as ocr_err:
                        print(f"OCR處理圖片時出錯: {ocr_err}")
                        result = None
                    
                    # 處理OCR結果
                    page_text = ""
                    if result is not None:
                        try:
                            for line in result:
                                if line is None:
                                    continue
                                for item in line:
                                    if len(item) >= 2 and isinstance(item[1], tuple) and len(item[1]) >= 1:
                                        # 獲取文本內容和置信度
                                        text_content = item[1][0]
                                        confidence = item[1][1]
                                        
                                        # 根據設置決定是否去除空白
                                        if remove_whitespace:
                                            text_content = text_content.replace(" ", "")
                                        
                                        page_text += text_content + "\n"
                        except TypeError as type_err:
                            print(f"處理OCR結果時出錯: {type_err}")
                        except Exception as proc_err:
                            print(f"處理OCR結果時出現未知錯誤: {proc_err}")
                    
                    ocr_text += f"===== 第{page_num+1}頁 =====\n{page_text}\n"
                    
                    # 刪除臨時圖片
                    try:
                        os.remove(img_path)
                    except Exception as e:
                        print(f"刪除臨時圖片時出錯: {e}")
        finally:
            # 清理臨時目錄
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"清理臨時目錄時出錯: {e}")
        
        # 不再在此處保存OCR結果，而是返回OCR文本，由調用者決定如何處理
        # save_txt參數保留以保持向後兼容性
        
        return ocr_text
    except Exception as e:
        print(f"使用PaddleOCR提取文本時出錯: {e}")
        return ""

def encrypt_pdf(input_path, output_path, user_pass, owner_pass, has_pikepdf=False, has_pypdf2=False):
    """加密PDF文件
    
    參數:
        input_path (str): 輸入文件路徑
        output_path (str): 輸出文件路徑
        user_pass (bytes): 用戶密碼
        owner_pass (bytes): 所有者密碼
        has_pikepdf (bool): 是否有pikepdf
        has_pypdf2 (bool): 是否有PyPDF2
        
    返回:
        bool: 是否成功
    """
    # 嘗試使用pikepdf加密
    if has_pikepdf:
        try:
            import pikepdf
            with pikepdf.open(input_path) as pdf:
                # 設置加密參數，使用AES-256加密
                encryption = pikepdf.Encryption(
                    user=user_pass.decode('utf-8') if isinstance(user_pass, bytes) else user_pass,
                    owner=owner_pass.decode('utf-8') if isinstance(owner_pass, bytes) else owner_pass,
                    allow=pikepdf.Permissions(extract=False, modify_annotation=False, modify_assembly=False,
                                             modify_form=False, modify_other=False, print_highres=False,
                                             print_lowres=False),
                    R=6,  # R=6 表示使用AES-256加密
                    aes=True
                )
                pdf.save(output_path, encryption=encryption)
            return True
        except Exception as e:
            print(f"使用pikepdf加密PDF時出錯: {e}")
    
    # 嘗試使用PyPDF2加密
    if has_pypdf2:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            # 複製所有頁面
            for page in reader.pages:
                writer.add_page(page)
            
            # 設置加密，使用AES-256加密
            writer.encrypt(
                user_password=user_pass.decode('utf-8') if isinstance(user_pass, bytes) else user_pass,
                owner_password=owner_pass.decode('utf-8') if isinstance(owner_pass, bytes) else owner_pass, algorithm="AES-256"
            )
            
            # 保存加密後的文件
            with open(output_path, "wb") as f:
                writer.write(f)
            return True
        except Exception as e:
            print(f"使用PyPDF2加密PDF時出錯: {e}")
    
    return False

def split_pdf(pdf_path, output_dir, pages_per_file, has_fitz=False, has_pikepdf=False, has_pypdf2=False):
    """分割PDF文件
    
    參數:
        pdf_path (str): PDF文件路徑
        output_dir (str): 輸出目錄
        pages_per_file (int): 每個文件的頁數
        has_fitz (bool): 是否有PyMuPDF
        has_pikepdf (bool): 是否有pikepdf
        has_pypdf2 (bool): 是否有PyPDF2
        
    返回:
        bool: 是否成功
    """
    split_success = False
    
    # 嘗試使用PyMuPDF分割
    if has_fitz and not split_success:
        try:
            import fitz
            with fitz.open(pdf_path) as pdf_doc:
                pdf_info = os.path.splitext(os.path.basename(pdf_path))[0]
                number_of_pages = pdf_doc.page_count
                
                print(f"PDF檔案...一共有{number_of_pages}頁，你最後會得到{(number_of_pages-1)//pages_per_file+1}個檔案")
                
                for i in range(0, number_of_pages, pages_per_file):
                    page_limit = min(i + pages_per_file, number_of_pages)
                    export_name = f"{pdf_info}-"
                    new_pdf_file_name = f"{export_name}{i//pages_per_file+1}.pdf"
                    page_file_path = os.path.join(output_dir, new_pdf_file_name)
                    
                    print(f"正在輸出...第{i+1}到{page_limit}頁，檔名：{new_pdf_file_name}")
                    
                    # 創建新的PDF文件並添加頁面
                    new_doc = None
                    try:
                        new_doc = fitz.open()
                        new_doc.insert_pdf(pdf_doc, from_page=i, to_page=page_limit-1)
                        new_doc.save(page_file_path)
                        print(f"使用PyMuPDF成功分割PDF：第{i+1}到{page_limit}頁")
                    except Exception as e:
                        print(f"使用PyMuPDF分割PDF時出錯: {e}")
                    finally:
                        if new_doc:
                            try:
                                new_doc.close()
                            except Exception as close_err:
                                print(f"關閉PDF文檔時出錯: {close_err}")
            split_success = True
        except Exception as e:
            print(f"使用PyMuPDF處理PDF時出錯: {e}")
    
    # 嘗試使用pikepdf分割
    if has_pikepdf and not split_success:
        try:
            import pikepdf
            with pikepdf.open(pdf_path) as pdf_doc:
                pdf_info = os.path.splitext(os.path.basename(pdf_path))[0]
                number_of_pages = len(pdf_doc.pages)
                
                print(f"PDF檔案...一共有{number_of_pages}頁，你最後會得到{(number_of_pages-1)//pages_per_file+1}個檔案")
                
                for i in range(0, number_of_pages, pages_per_file):
                    page_limit = min(i + pages_per_file, number_of_pages)
                    export_name = f"{pdf_info}-"
                    new_pdf_file_name = f"{export_name}{i//pages_per_file+1}.pdf"
                    page_file_path = os.path.join(output_dir, new_pdf_file_name)
                    
                    print(f"正在輸出...第{i+1}到{page_limit}頁，檔名：{new_pdf_file_name}")
                    
                    # 創建新的PDF文件並添加頁面
                    new_pdf = None
                    try:
                        new_pdf = pikepdf.new()
                        for j in range(i, page_limit):
                            new_pdf.pages.append(pdf_doc.pages[j])
                        
                        # 保存新文件
                        new_pdf.save(page_file_path)
                    except Exception as e:
                        print(f"使用pikepdf創建分割文件時出錯: {e}")
                    finally:
                        if new_pdf is not None and hasattr(new_pdf, 'close'):
                            try:
                                new_pdf.close()
                            except Exception as close_err:
                                print(f"關閉pikepdf文檔時出錯: {close_err}")
            split_success = True
        except Exception as e:
            print(f"使用pikepdf分割PDF時出錯: {e}")
    
    # 嘗試使用PyPDF2分割
    if has_pypdf2 and not split_success:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            try:
                from tqdm import tqdm
            except ImportError:
                def tqdm(iterable, **kwargs):
                    return iterable
                    
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                number_of_pages = len(reader.pages)
                
                print(f"PDF檔案...一共有{number_of_pages}頁，你最後會得到{(number_of_pages-1)//pages_per_file+1}個檔案")
                
                for i in tqdm(range(0, number_of_pages, pages_per_file), desc="分割PDF進度"):
                    page_limit = min(i + pages_per_file, number_of_pages)
                    pdf_info = os.path.splitext(os.path.basename(pdf_path))[0]
                    export_name = f"{pdf_info}-"
                    new_pdf_file_name = f"{export_name}{i//pages_per_file+1}.pdf"
                    page_file_path = os.path.join(output_dir, new_pdf_file_name)
                    
                    print(f"正在輸出...第{i+1}到{page_limit}頁，檔名：{new_pdf_file_name}")
                    
                    writer = PdfWriter()
                    output_file = None
                    try:
                        for j in range(i, page_limit):
                            writer.add_page(reader.pages[j])
                        
                        with open(page_file_path, 'wb') as output_file:
                            writer.write(output_file)
                    except Exception as writer_error:
                        print(f"使用PyPDF2寫入頁面時出錯: {writer_error}")
            split_success = True
        except Exception as e:
            print(f"使用PyPDF2分割PDF時出錯: {e}")
    
    return split_success

def process_pdf_files(pdf_files, rule_items, search_location, ori_meta, has_fitz, has_pypdf2, has_paddleocr, has_pikepdf, max_workers=4, is_copy_mode=False, use_ocr=False, remove_whitespace=False, save_ocr_txt=False, default_user_password=None, default_owner_password=None):
    """處理PDF文件
    
    參數:
        pdf_files (list): PDF文件列表
        rule_items (list): 規則列表
        search_location (str): 搜索位置
        ori_meta (bool): 是否保留原始元數據
        has_fitz (bool): 是否有PyMuPDF
        has_pypdf2 (bool): 是否有PyPDF2
        has_paddleocr (bool): 是否有PaddleOCR
        has_pikepdf (bool): 是否有pikepdf
        max_workers (int): 最大工作線程數
        is_copy_mode (bool): 是否為複製模式（True為複製，False為重命名）
        use_ocr (bool): 是否使用OCR功能處理所有PDF文件
        remove_whitespace (bool): 是否去除中文文本中的空白
        save_ocr_txt (bool): 是否將OCR結果保存為txt文件
        default_user_password (str): 默認用戶密碼
        default_owner_password (str): 默認所有者密碼
        
    返回:
        int: 處理的文件數量
    """
    # 記錄開始時間
    start_time = time.time()
    
    if not pdf_files:
        log_message(f"在 {search_location} 中未找到PDF文件", level='警告')
        return 0
    
    log_message(f"找到 {len(pdf_files)} 個PDF文件")
    
    # 如果啟用OCR但沒有PaddleOCR，顯示警告
    if use_ocr and not has_paddleocr:
        log_message("警告：已啟用OCR功能，但未安裝PaddleOCR。將使用標準文本提取方法。", level='警告')
    elif use_ocr and has_paddleocr:
        log_message("已啟用OCR功能，將使用PaddleOCR處理PDF文件", level='信息')
        if remove_whitespace:
            log_message("已啟用去除中文空白功能", level='信息')
        if save_ocr_txt:
            log_message("已啟用保存OCR結果為txt文件功能", level='信息')
    
    # 導入file_renamer函數
    from file_utils import file_renamer
    
    # 定義單個PDF處理函數
    def process_single_pdf(pdf_file):
        try:
            # 使用file_renamer函數處理PDF文件
            return file_renamer(
                rule_items=rule_items,
                pdf_file=pdf_file,
                search_location=search_location,
                is_copy_mode=is_copy_mode,
                default_user_password=default_user_password,
                default_owner_password=default_owner_password,
                has_fitz=has_fitz,
                has_pypdf2=has_pypdf2,
                has_paddleocr=has_paddleocr,
                has_pikepdf=has_pikepdf,
                use_ocr=use_ocr,
                remove_whitespace=remove_whitespace,
                save_ocr_txt=save_ocr_txt
            )
        except Exception as e:
            log_message(f"處理文件時出錯: {pdf_file}, {e}", level='错误')
            return False
    
    # 使用並行處理函數處理所有PDF文件
    from main import process_files_parallel
    processed_count = process_files_parallel(pdf_files, process_single_pdf, max_workers)
    
    # 計算總運行時間
    end_time = time.time()
    total_time = end_time - start_time
    hours, remainder = divmod(total_time, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    # 顯示總運行時間
    time_str = ""
    if hours > 0:
        time_str += f"{int(hours)}小時"
    if minutes > 0 or hours > 0:
        time_str += f"{int(minutes)}分"
    time_str += f"{seconds:.2f}秒"
    
    log_message(f"PDF處理完成！總共處理了{processed_count}個文件，耗時{time_str}", level='信息')
    print(f"\n總共處理了{processed_count}個文件，耗時{time_str}")
    
    return processed_count