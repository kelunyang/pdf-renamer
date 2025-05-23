import sys
import os

# 嘗試導入questionary，如果不可用則在check_and_install_dependencies中安裝
try:
    import questionary
except ImportError:
    questionary = None

# 全局變量，將在main中初始化
default_user_password = ""
default_owner_password = ""

def input_helper(prompt, nullable=False, default="", validation_func=None, is_password=False, is_confirm=False, confirm_default_answer=True):
    """
    增強的輸入處理函數
    
    參數:
        prompt (str): 提示用戶的信息
        nullable (bool): 是否允許空值 (ignored if is_confirm=True)
        default (str): 如果允許空值時輸入為空，則使用此默認值 (for text input, ignored if is_confirm=True)
        validation_func (callable): 可選的驗證函數 (ignored if is_confirm=True)
        is_password (bool): 是否為密碼輸入 (ignored if is_confirm=True)
        is_confirm (bool): 是否為確認型問題 (yes/no)
        confirm_default_answer (bool): 確認型問題的默認答案 (True for Yes, False for No)
        
    返回:
        str or bool: 用戶輸入 (bool if is_confirm=True, otherwise str)
    """
    # 檢查questionary是否可用
    global questionary
    if questionary is None:
        print("questionary模組未安裝，嘗試安裝...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "questionary"])
            import questionary
            print("questionary安裝成功!")
        except Exception as e:
            print(f"安裝questionary時出錯: {e}")
            print("將使用標準輸入代替。")
    
    if is_confirm:
        # 使用questionary進行確認型問題
        if questionary is not None:
            return questionary.confirm(prompt, default=confirm_default_answer).ask()
        else:
            # 如果questionary不可用，使用標準輸入
            print(f"{prompt} (y/n) [{'y' if confirm_default_answer else 'n'}]: ")
            response = input().strip().lower()
            if not response:
                return confirm_default_answer
            return response in ['y', 'yes']

    # Existing logic for non-confirm questions:
    attempts = 0
    max_attempts = 3  # 最大嘗試次數，防止無限循環
    
    # 如果是密碼輸入，且沒有提供默認值，則使用全局隨機密碼
    if is_password and not default: # This 'default' is the string default for text input
        if "開啟" in prompt:
            global default_user_password
            actual_default = default_user_password
        elif "編輯" in prompt:
            global default_owner_password
            actual_default = default_owner_password
        else:
            actual_default = default # The original string default
    else:
        actual_default = default # The original string default
    
    while attempts < max_attempts:
        # 構建提示信息
        display_prompt = prompt
        
        if nullable:
            if actual_default:
                if is_password:
                    display_prompt += f" [按Enter使用預設密碼: {actual_default}]"
                else:
                    display_prompt += f" [按Enter使用默認值: {actual_default}]"
            else:
                display_prompt += " [按Enter跳過]"
        else:
            display_prompt += " [必須輸入]"
            
        if attempts > 0:
            display_prompt += " [請再試一次]"
            
        # 獲取用戶輸入
        # For non-confirm questions, we still use print and input
        print(display_prompt) # Print the prompt
        value = input().strip() # Get text input
        
        # 處理空輸入
        if not value:
            if nullable:
                return actual_default
            else:
                attempts += 1
                print("此項不可為空，請輸入值。")
                continue
        
        # 執行驗證函數（如果提供）
        if validation_func:
            is_valid, error_msg = validation_func(value)
            if not is_valid:
                print(f"輸入無效: {error_msg}")
                attempts += 1
                continue
        
        return value # Return string value
    
    # 如果達到最大嘗試次數
    if nullable:
        print(f"已達最大嘗試次數，使用默認值: {actual_default}")
        return actual_default
    else:
        print("已達最大嘗試次數，請重新運行程式並提供有效輸入。")
        sys.exit(0)

def validate_path(path):
    """
    驗證路徑是否存在
    
    參數:
        path (str): 要驗證的路徑
        
    返回:
        tuple: (是否有效, 錯誤消息)
    """
    if not path:
        return False, '路徑不可為空'
    if not os.path.exists(path):
        return False, '路徑不存在'
    return True, ''