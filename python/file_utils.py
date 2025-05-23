import os
import gc
import time
import sys
import signal
import importlib.util
import subprocess
import threading
from log_utils import log_message

# 全局變量，用於標記是否收到中斷信號
interrupt_received = False

def is_file_in_use(file_path):
    """
    跨平台檢測文件是否被其他程序占用
    
    參數:
        file_path: 文件路徑
        
    返回:
        bool: 如果文件被占用返回True，否則返回False
    """
    if not os.path.exists(file_path):
        return False
        
    try:
        # 嘗試以寫入模式打開文件
        # 如果文件被占用，這將引發異常
        with open(file_path, 'a+b') as f:
            # 嘗試獲取獨占鎖
            try:
                # 不同平台的文件鎖定方法
                if sys.platform.startswith('win'):
                    # Windows平台
                    try:
                        import msvcrt
                        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                    except ImportError:
                        # 如果msvcrt不可用，嘗試使用win32file
                        try:
                            import win32file
                            hfile = win32file._get_osfhandle(f.fileno())
                            win32file.LockFileEx(hfile, win32file.LOCKFILE_EXCLUSIVE_LOCK | win32file.LOCKFILE_FAIL_IMMEDIATELY, 0, 1, 0, None)
                            win32file.UnlockFileEx(hfile, 0, 1, 0, None)
                        except ImportError:
                            # 如果win32file也不可用，使用基本的文件操作測試
                            pass
                else:
                    # Unix/Linux/Mac平台
                    try:
                        import fcntl
                        fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.flock(f.fileno(), fcntl.LOCK_UN)
                    except ImportError:
                        # 如果fcntl不可用，使用基本的文件操作測試
                        pass
                        
                # 如果能夠執行到這裡，說明文件沒有被占用
                return False
            except (IOError, OSError):
                # 獲取鎖失敗，文件被占用
                return True
    except (IOError, OSError):
        # 無法打開文件，文件被占用
        return True
        
    # 默認情況下，假設文件沒有被占用
    return False

def file_utils_signal_handler(sig, frame):
    """處理Ctrl+C信號，設置中斷標記"""
    global interrupt_received
    interrupt_received = True
    print("\n\n收到中斷信號，將在當前任務完成後退出...")
    log_message("收到中斷信號，將在當前任務完成後退出...", level='警告')

def file_renamer(rule_items, pdf_file, search_location, result_queue=None, ui_update_event=None, is_copy_mode=False, default_user_password=None, default_owner_password=None, has_fitz=False, has_pypdf2=False, has_paddleocr=False, has_pikepdf=False, use_ocr=False, remove_whitespace=False, save_ocr_txt=False):
    """
    根據規則匹配PDF內容並重命名或複製PDF文件，如需要還會加密
    
    參數:
        rule_items: 規則項目列表
        pdf_file: PDF文件路徑
        search_location: 搜索位置
        result_queue: 結果隊列
        ui_update_event: UI更新事件 (已棄用，保留參數以兼容現有代碼)
        is_copy_mode: 是否為複製模式（True為複製，False為重命名）
        default_user_password: 默認用戶密碼
        default_owner_password: 默認所有者密碼
        has_fitz: 是否有PyMuPDF
        has_pypdf2: 是否有PyPDF2
        has_paddleocr: 是否有PaddleOCR
        has_pikepdf: 是否有pikepdf
        use_ocr: 是否使用OCR功能處理所有PDF文件
        remove_whitespace: 是否去除中文文本中的空白
        save_ocr_txt: 是否將OCR結果保存為txt文件
    """
    
    # 如果result_queue未傳入且全局變量中沒有定義，創建一個新的隊列
    if result_queue is None and 'result_queue' not in globals():
        import queue
        result_queue = queue.Queue()
    
    # 導入需要的模塊
    import re
    from pdf_utils import extract_text_from_pdf, encrypt_pdf
    
    # 檢查是否收到中斷信號
    if interrupt_received:
        log_message(f"由於收到中斷信號，跳過處理文件: {pdf_file}", level='警告')
        if result_queue:
            result_queue.put((pdf_file, False, None))
        return False
        
    try:
        # 提取文件名（不含路徑和擴展名）
        filename = os.path.splitext(os.path.basename(pdf_file))[0]
        
        # 如果需要OCR，為每個線程創建獨立的PaddleOCR實例
        ocr_instance = None
        if has_paddleocr and use_ocr:
            try:
                # 使用paddle_utils模塊初始化PaddleOCR
                try:
                    from paddle_utils import init_paddleocr
                    ocr_instance = init_paddleocr(use_angle_cls=True, lang="ch")
                    if ocr_instance is None:
                        # 如果初始化失敗，嘗試直接導入PaddleOCR
                        from paddleocr import PaddleOCR
                        ocr_instance = PaddleOCR(use_angle_cls=True, lang="ch")
                except ImportError:
                    # 如果paddle_utils不可用，直接導入PaddleOCR
                    from paddleocr import PaddleOCR
                    ocr_instance = PaddleOCR(use_angle_cls=True, lang="ch")
                log_message(f"為線程創建獨立的PaddleOCR實例處理文件: {filename}", level='信息')
            except Exception as e:
                log_message(f"創建PaddleOCR實例時出錯: {e}", level='警告')
        
        # 提取元數據
        metadata = ""
        if has_fitz:
            try:
                import fitz
                with fitz.open(pdf_file) as doc:
                    if doc.metadata and isinstance(doc.metadata, dict):
                        for key, value in doc.metadata.items():
                            if value:
                                metadata += str(value)
            except Exception as e:
                log_message(f"讀取元數據時出錯: {e}", level='警告')
        
        # 初始化變量，用於跟踪是否成功重命名
        rename_success = False
        new_pdf_path = None
        
        # 應用規則
        for rule in rule_items:
            # 檢查是否收到中斷信號
            if interrupt_received:
                log_message(f"由於收到中斷信號，中止規則處理: {pdf_file}", level='警告')
                if result_queue:
                    result_queue.put((pdf_file, False, None))
                return False
            # 提取文本，如果use_ocr為True則強制使用OCR，但暫時不保存OCR結果
            text = extract_text_from_pdf(pdf_file, has_fitz, has_pypdf2, has_paddleocr, force_ocr=use_ocr, remove_whitespace=remove_whitespace, save_txt=False, ocr_instance=ocr_instance)
            
            # 根據目標類型選擇要匹配的內容
            if rule.target_type == "內容":
                content_to_match = text
            elif rule.target_type == "檔名":
                content_to_match = filename
            elif rule.target_type == "元數據":
                content_to_match = metadata
            else:
                content_to_match = text  # 默認使用內容
            
            # 使用正則表達式匹配
            pattern = re.compile(rule.pattern) if hasattr(rule, 'pattern') else rule.rule_from
            matches = pattern.findall(content_to_match)
            occurrence = int(rule.occurrence) if hasattr(rule, 'occurrence') else rule.occurrence_match
            
            # 如果匹配成功
            if matches and len(matches) >= occurrence:
                # 獲取新文件名
                new_name = rule.name if hasattr(rule, 'name') else rule.name_to
                
                # 構建新文件路徑
                output_path = os.path.join(os.path.dirname(pdf_file), f"{new_name}.pdf")
                
                # 處理文件名衝突
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(os.path.dirname(pdf_file), f"{new_name}_{counter}.pdf")
                    counter += 1
                    if counter > 100: # 防止無限循環，設定一個上限
                        log_message(f"嘗試了100個不同的檔名後，仍然無法為 {new_name} 找到唯一的檔名。跳過此檔案。", level='警告')
                        if result_queue:
                            result_queue.put((pdf_file, False, None))
                        return False
                
                # 如果需要加密
                # 如果需要加密
                if hasattr(rule, 'user_pass') and hasattr(rule, 'owner_pass'):
                    # 設置密碼
                    if hasattr(rule, 'user_pass_set') and rule.user_pass_set:
                        # 使用函數參數中傳入的默認用戶密碼
                        user_pass = default_user_password
                        # 如果密碼為空，生成一個隨機密碼
                        if not user_pass:
                            import random, string
                            user_pass = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))
                            print(f"警告：用戶密碼為空，已生成隨機密碼: {user_pass}")
                    else:
                        # 如果rule中有user_pass屬性且不為空，則使用rule中的密碼
                        if hasattr(rule, 'user_pass') and rule.user_pass != '':
                            user_pass = rule.user_pass
                        else:
                            # 否則使用默認密碼
                            user_pass = default_user_password
                            # 如果默認密碼為空，生成一個隨機密碼
                            if not user_pass:
                                import random, string
                                user_pass = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))
                                print(f"警告：用戶密碼為空，已生成隨機密碼: {user_pass}")
                    
                    if hasattr(rule, 'owner_pass_set') and rule.owner_pass_set:
                        # 使用函數參數中傳入的默認所有者密碼
                        owner_pass = default_owner_password
                        # 如果密碼為空，生成一個隨機密碼
                        if not owner_pass:
                            import random, string
                            owner_pass = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))
                            print(f"警告：所有者密碼為空，已生成隨機密碼: {owner_pass}")
                    else:
                        # 如果rule中有owner_pass屬性且不為空，則使用rule中的密碼
                        if hasattr(rule, 'owner_pass') and rule.owner_pass != '':
                            owner_pass = rule.owner_pass
                        else:
                            # 否則使用默認密碼
                            owner_pass = default_owner_password
                            # 如果默認密碼為空，生成一個隨機密碼
                            if not owner_pass:
                                import random, string
                                owner_pass = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))
                                print(f"警告：所有者密碼為空，已生成隨機密碼: {owner_pass}")
                    
                    # 使用臨時文件路徑進行加密
                    temp_output_path = output_path + ".temp"
                    encrypt_success = encrypt_pdf(
                        pdf_file, 
                        temp_output_path, 
                        user_pass, 
                        owner_pass, 
                        has_pikepdf, 
                        has_pypdf2
                    )
                    
                    if encrypt_success:
                        try:
                            # 重命名源文件為備份（以防萬一）
                            backup_path = pdf_file + ".bak"
                            os.rename(pdf_file, backup_path)
                            
                            # 將臨時文件重命名為最終文件名
                            os.rename(temp_output_path, output_path)
                            
                            # 刪除備份文件
                            try:
                                os.remove(backup_path)
                            except Exception as backup_err:
                                log_message(f"刪除備份文件時出錯: {backup_err}", level='警告')
                            
                            log_message(f"文件已加密並重命名為: {output_path}")
                            rename_success = True
                            new_pdf_path = output_path
                        except Exception as rename_err:
                            log_message(f"重命名加密文件時出錯: {rename_err}", level='错误')
                            rename_success = False
                    else:
                        log_message(f"加密文件失敗: {pdf_file}", level='警告')
                        rename_success = False
                else:
                    # 根據模式選擇重命名或複製
                    try:
                        # 檢查文件是否被占用
                        file_in_use = is_file_in_use(pdf_file)
                        
                        # 如果文件被占用且不是複製模式，自動切換到複製模式
                        if file_in_use and not is_copy_mode:
                            log_message(f"檔案 {pdf_file} 被其他程序占用，自動切換到複製模式", level='警告')
                            is_copy_mode = True
                        
                        if is_copy_mode:
                            # 複製模式：複製文件
                            import shutil
                            shutil.copy2(pdf_file, output_path)
                            log_message(f"文件已複製為: {output_path}")
                            rename_success = True
                            new_pdf_path = output_path
                        else:
                            # 重命名模式：嘗試直接重命名
                            os.rename(pdf_file, output_path)
                            log_message(f"文件已重命名為: {output_path}")
                            rename_success = True
                            new_pdf_path = output_path
                    except Exception as e:
                        # 如果重命名失敗（可能是跨卷），嘗試複製後刪除
                        try:
                            import shutil
                            shutil.copy2(pdf_file, output_path)
                            if not is_copy_mode:  # 只有在重命名模式下才刪除原文件
                                try:
                                    os.remove(pdf_file)
                                    log_message(f"文件已複製並刪除原文件: {output_path}")
                                except Exception as del_err:
                                    log_message(f"無法刪除原文件，可能被占用: {del_err}", level='警告')
                                    log_message(f"文件已複製為: {output_path}")
                            else:
                                log_message(f"文件已複製為: {output_path}")
                            rename_success = True
                            new_pdf_path = output_path
                        except Exception as copy_err:
                            log_message(f"重命名/複製文件失敗: {e}, {copy_err}", level='错误')
                            rename_success = False
                
                # 找到匹配後跳出規則循環，處理下一個文件
                break
        
        # 如果沒有匹配的規則
        if not rename_success:
            log_message(f"沒有匹配的規則或處理失敗: {pdf_file}", level='警告')
        
        # 無論是否重命名成功，只要啟用了OCR和保存OCR結果，都將OCR文本保存到txt文件中
        if use_ocr and save_ocr_txt and has_paddleocr and text:
            # 如果重命名成功，使用新的文件路徑；否則使用原始文件路徑
            output_path = new_pdf_path if new_pdf_path else pdf_file
            txt_path = os.path.splitext(output_path)[0] + '_ocr.txt'
            try:
                with open(txt_path, 'w', encoding='utf-8-sig') as f:
                    f.write(text)
                log_message(f"已保存OCR結果到: {txt_path}", level='信息')
            except Exception as txt_err:
                log_message(f"保存OCR結果到文件時出錯: {txt_err}", level='警告')

        
        # 將結果放入隊列
        if result_queue:
            result_queue.put((pdf_file, rename_success, new_pdf_path))
        
        return rename_success
    except Exception as e:
        log_message(f"處理文件時出錯: {pdf_file}, {e}", level='错误')
        if result_queue:
            result_queue.put((pdf_file, False, None))
        return False


            


        


def check_and_install_dependencies():
    """檢查並安裝所需的依賴庫"""
    # 注意：不要在此处导入第三方库
    import site # Add site import here
    import shutil # 用於檢查ccache是否存在
    
    # 檢查是否在虛擬環境中運行
    in_virtualenv = hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)
    
    # 如果在虛擬環境中運行，確保虛擬環境的site-packages目錄在sys.path中
    if in_virtualenv:
        print("檢測到Python虛擬環境，正在確保路徑設置正確...")
        venv_site_packages = site.getsitepackages()[0]
        if venv_site_packages not in sys.path:
            sys.path.insert(0, venv_site_packages)
            print(f"已將虛擬環境路徑 {venv_site_packages} 添加到sys.path")
        
        # 刷新importlib緩存以確保能找到新安裝的包
        importlib.invalidate_caches()
        
    # 檢查是否安裝了ccache（用於加速PaddleOCR的編譯）
    ccache_installed = shutil.which('ccache') is not None
    if not ccache_installed:
        print("警告: 未檢測到ccache。PaddleOCR可能需要重新編譯所有源文件，這會導致首次運行較慢。")
        print("您可以從以下網址下載並安裝ccache: https://github.com/ccache/ccache/blob/master/doc/INSTALL.md")
        print("Windows用戶可以從這裡下載預編譯版本: https://github.com/ccache/ccache/releases")
    else:
        print("已檢測到ccache，PaddleOCR編譯將更快速。")
    required_packages = {
        "PyPDF2": "PyPDF2",
        "pikepdf": "pikepdf",
        "fitz": "PyMuPDF",  # fitz是PyMuPDF的一部分
        "win32file": "pywin32", # 用於Windows特定文件操作
        "psutil": "psutil", # 用於進程管理和檢測
        "paddleocr": "paddleocr", # 用於OCR識別
        "paddle": "paddlepaddle", # PaddleOCR的依賴
        "questionary": "questionary", # 添加 questionary
        "tqdm": "tqdm", # 添加 tqdm
        "re": None,  # 標準庫不需要安裝
        "csv": None,  # 標準庫不需要安裝
        "gc": None    # 標準庫不需要安裝
    }
    
    # 添加備用庫以防主要庫安裝失敗
    alternative_packages = {
        "PyMuPDF": ["pymupdf"],
        "pikepdf": ["pdfrw"],  # 如果pikepdf安裝失敗，可以嘗試pdfrw
        "PyPDF2": ["pdfrw"],    # 如果PyPDF2安裝失敗，也可以嘗試pdfrw
        "questionary": []  # questionary沒有備用庫，必須安裝
    }
    
    missing_packages = []
    
    print("檢查必要的Python庫...")
    
    for module_name, package_name in required_packages.items():
        if package_name is None:  # 標準庫，不需要檢查
            continue
            
        if importlib.util.find_spec(module_name) is None:
            missing_packages.append(package_name)
    
    if missing_packages:
        print(f"檢測到缺少以下庫: {', '.join(missing_packages)}")
        
        # 詢問是否自動安裝
        try:
            import questionary
            install_choice = questionary.select(
                "是否要自動安裝這些庫？",
                choices=["是", "否"],
                default="是"
            ).ask() == "是"
        except ImportError:
            print("是否要自動安裝這些庫？(y/n) [默认y]: ")
            install_input = input().strip().lower()
            install_choice = install_input in ('y', '')
        
        if install_choice:
            print("正在安裝缺少的庫...")
            
            # 嘗試以用戶模式安裝包（避免許可權問題）
            for package in missing_packages:
                print(f"安裝 {package}...")
                try:
                    # 使用 --user 標誌以用戶模式安裝
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])
                    print(f"{package} 安裝成功!")
                except subprocess.CalledProcessError as e:
                    print(f"以用戶模式安裝 {package} 時出錯: {e}")
                    print("嘗試備用安裝方法...")
                    
                    # 嘗試以管理員權限安裝（僅適用於Windows）
                    try_admin = False
                    if sys.platform.startswith('win'):
                        try:
                            import questionary
                            try_admin = questionary.confirm(f"嘗試以管理員權限安裝 {package}？", default=False).ask()
                        except ImportError:
                            print(f"嘗試以管理員權限安裝 {package}？(y/n) [默认n]: ")
                            try_admin_input = input().strip().lower()
                            try_admin = try_admin_input == 'y'
                        # try_admin will be True or False directly
                    
                    if try_admin:
                        try:
                            # 以管理員權限運行安裝命令
                            admin_cmd = f'runas /user:Administrator "pip install {package}"'
                            os.system(admin_cmd)
                            
                            # 檢查安裝是否成功
                            if importlib.util.find_spec(module_name) is not None:
                                print(f"{package} 安裝成功!")
                                continue
                            else:
                                print(f"無法確認 {package} 是否安裝成功，嘗試備用庫...")
                        except Exception as admin_error:
                            print(f"以管理員權限安裝失敗: {admin_error}")
                    
                    # 嘗試安裝備用庫
                    if package in alternative_packages:
                        for alt_package in alternative_packages[package]:
                            try:
                                print(f"嘗試安裝備用庫 {alt_package}...")
                                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", alt_package])
                                print(f"備用庫 {alt_package} 安裝成功!")
                                
                                # 為確保函數能使用新的庫，修改importlib搜索路徑
                                if package == "PyMuPDF" and alt_package == "pymupdf":
                                    sys.modules["fitz"] = importlib.import_module("fitz")
                                break
                            except subprocess.CalledProcessError:
                                print(f"安裝備用庫 {alt_package} 失敗")
                    
                    print("自動安裝失敗，請手動運行以下命令（可能需要管理員權限）:")
                    print(f"pip install {package}")
                    print("或者使用虛擬環境:")
                    print("python -m venv pdf_env")
                    print("pdf_env\\Scripts\\activate  # Windows")
                    print(f"pip install {package}")
                    
                    # 詢問是否要繼續而不使用此庫
                    try:
                        import questionary
                        continue_choice = questionary.select(
                            "是否要繼續而不使用此庫？這可能導致功能受限",
                            choices=["是", "否"],
                            default="否"
                        ).ask() == "是"
                    except ImportError:
                        print("是否要繼續而不使用此庫？這可能導致功能受限(y/n) [默认n]: ")
                        continue_input = input().strip().lower()
                        continue_choice = continue_input == 'y'
                    if not continue_choice: # if user chooses No (False)
                        sys.exit(1)
            
            print("所有可安裝的庫已安裝完成，重新載入程式...")
            
            # 重新載入模組
            for module_name in required_packages.keys():
                if module_name not in ['re', 'csv', 'gc', 'win32file']:
                    try:
                        if module_name in sys.modules:
                            importlib.reload(sys.modules[module_name])
                        else:
                            # 如果模組尚未載入，則導入
                            try:
                                globals()[module_name] = importlib.import_module(module_name)
                            except ImportError:
                                pass # 允許 win32file 在非 Windows 環境下導入失敗
                    except Exception as reload_error:
                        print(f"重新載入 {module_name} 時出錯: {reload_error}")
                elif module_name == 'win32file' and sys.platform.startswith('win'):
                    try:
                        importlib.invalidate_caches() # Invalidate caches first
                        if module_name in sys.modules:
                            # If already imported, try to reload
                            importlib.reload(sys.modules[module_name])
                        else:
                            # If not imported, try to import
                            globals()[module_name] = importlib.import_module(module_name)
                        print(f"{module_name} 重新載入成功。")
                    except ImportError: # Specifically catch ImportError if the module is not found
                        print(f"重新載入 {module_name} 時出錯 (ImportError)，嘗試添加 user site-packages 到 sys.path...")
                        # Attempt to add user site-packages to sys.path
                        user_site_packages = site.USER_SITE
                        if user_site_packages and user_site_packages not in sys.path:
                            sys.path.append(user_site_packages)
                            print(f"已將 {user_site_packages} 添加到 sys.path。")
                            # Invalidate caches again after modifying sys.path
                            importlib.invalidate_caches()
                        
                        # Try importing again
                        try:
                            # It's crucial to ensure the module is freshly imported
                            if module_name in sys.modules:
                                del sys.modules[module_name] # Remove potentially problematic or old entry
                            globals()[module_name] = importlib.import_module(module_name)
                            print(f"{module_name} 在添加路徑後成功導入。")
                        except Exception as inner_import_error:
                            print(f"在添加 user site-packages 到 sys.path 後，導入 {module_name} 仍然失敗: {inner_import_error}")
                    except Exception as reload_error:
                         print(f"重新載入 {module_name} 時發生其他錯誤: {reload_error}")
        else:
            print("您選擇不安裝缺少的庫。程式將嘗試以有限功能運行。")
            print("如果發生錯誤，請手動安裝以下庫後再運行程式:")
            for package in missing_packages:
                print(f"pip install {package}")
    else:
        print("所有必要的庫已安裝!")
    
    # 檢查是否有PaddleOCR
    has_paddleocr = importlib.util.find_spec("paddleocr") is not None and importlib.util.find_spec("paddle") is not None
    if not has_paddleocr:
        print("未檢測到PaddleOCR，無法使用OCR功能")
        # 詢問用戶是否要安裝PaddleOCR
        try:
            import questionary
            install_paddleocr = questionary.select(
                "是否要安裝PaddleOCR以啟用OCR功能？",
                choices=["是", "否"],
                default="否"
            ).ask() == "是"
        except ImportError:
            install_paddleocr = input("是否要安裝PaddleOCR以啟用OCR功能？(y/n): ").strip().lower() in ['y', 'yes']
        
        if install_paddleocr:
            try:
                print("正在安裝PaddleOCR相關套件，這可能需要一些時間...")
                # 先安裝paddlepaddle（CPU版本，較小）
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "paddlepaddle"])
                print("PaddlePaddle安裝成功")
                # 再安裝paddleocr
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "paddleocr"])
                print("PaddleOCR安裝成功，現在可以使用OCR功能了")
                # 更新狀態
                has_paddleocr = True
            except Exception as e:
                print(f"安裝PaddleOCR時出錯: {e}")
                print("您可以稍後手動安裝: pip install --user paddlepaddle paddleocr")
    else:
        print("已檢測到PaddleOCR，可以使用OCR功能")
        
    # 檢查PyTorch的shm.dll問題
    if sys.platform.startswith('win'):
        # 檢查是否安裝了PyTorch
        has_torch = importlib.util.find_spec("torch") is not None
        if has_torch:
            try:
                import torch
                torch_lib_path = os.path.join(os.path.dirname(torch.__file__), 'lib')
                shm_dll_path = os.path.join(torch_lib_path, 'shm.dll')
                
                if os.path.exists(shm_dll_path):
                    print(f"檢測到PyTorch的shm.dll位於: {shm_dll_path}")
                    # 檢查環境變數PATH中是否包含torch/lib目錄
                    env_path = os.environ.get('PATH', '')
                    if torch_lib_path not in env_path:
                        print(f"將PyTorch的lib目錄添加到PATH環境變數: {torch_lib_path}")
                        os.environ['PATH'] = torch_lib_path + os.pathsep + env_path
                        print("已臨時添加到PATH環境變數，如需永久添加，請修改系統環境變數設置")
                else:
                    print("警告: 未找到PyTorch的shm.dll文件，可能會影響PaddleOCR的運行")
            except Exception as e:
                print(f"檢查PyTorch的shm.dll時出錯: {e}")
        else:
            print("未檢測到PyTorch，如果PaddleOCR運行時出現shm.dll錯誤，可能需要安裝PyTorch")
            print("您可以使用以下命令安裝PyTorch: pip install --user torch torchvision torchaudio")
    
    return True