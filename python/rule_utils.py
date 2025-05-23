import re
from log_utils import log_message

class Rule:
    def __init__(self, rule_pattern, name, target_type, occurrence, user_pass, owner_pass, user_pass_set, owner_pass_set, encrypt_enable):
        # 聲明全局變量，必須在使用前聲明
        from input_utils import default_user_password, default_owner_password
        
        try:
            # 處理正則表達式模式，確保它是有效的
            self.rule_from = re.compile(rule_pattern, re.DOTALL)
            self.name_to = name
            self.target_type = target_type
            
            # 確保occurrence_match是整數
            if isinstance(occurrence, int):
                self.occurrence_match = occurrence
            else:
                try:
                    self.occurrence_match = int(occurrence)
                except ValueError:
                    print(f"警告: 將'{occurrence}'轉換為整數失敗，設置為1")
                    self.occurrence_match = 1
            
            # 處理密碼 - 如果用戶沒有輸入密碼，使用全局默認密碼
            # 全局默認密碼在程序啟動時生成，存儲在全局變量中
            self.user_pass = user_pass if user_pass else default_user_password
            self.owner_pass = owner_pass if owner_pass else default_owner_password
            
            # 將密碼轉換為bytes格式
            self.user_pass = self.user_pass.encode('utf-8')
            self.owner_pass = self.owner_pass.encode('utf-8')
            
            user_prompt = "開啟密碼已設定" if not user_pass_set else "開啟密碼採用預設密碼"
            owner_prompt = "編輯密碼已設定" if not owner_pass_set else "編輯密碼採用預設密碼"
            pass_prompt = f"（{user_prompt}／{owner_prompt}）" if encrypt_enable else ""
            
            # 將訊息寫入日誌而不是直接打印
            log_message(f"找：{rule_pattern}的{target_type}，重複出現{self.occurrence_match}次，更名為：{name}{pass_prompt}", level='信息')
        except re.error as e:
            log_message(f"警告: 正則表達式'{rule_pattern}'無效: {e}", level='警告')
            # 設置一個永不匹配的默認正則表達式
            self.rule_from = re.compile(r'a^')  # 這個正則表達式永遠不會匹配任何內容
            self.name_to = name
            self.target_type = target_type
            self.occurrence_match = 0
            
            # 使用全局默認密碼
            self.user_pass = (user_pass if user_pass else default_user_password).encode('utf-8')
            self.owner_pass = (owner_pass if owner_pass else default_owner_password).encode('utf-8')
            
            log_message(f"規則將被忽略: {rule_pattern}", level='警告')
        except Exception as e:
            log_message(f"創建規則時出錯: {e}", level='错误')
            # 設置安全的默認值
            self.rule_from = re.compile(r'a^')
            self.name_to = name
            self.target_type = target_type
            self.occurrence_match = 0
            
            # 使用全局默認密碼
            self.user_pass = (user_pass if user_pass else default_user_password).encode('utf-8')
            self.owner_pass = (owner_pass if owner_pass else default_owner_password).encode('utf-8')

# 簡化版的Rule類，用於向後兼容
class SimpleRule:
    def __init__(self, pattern, replacement, priority, rule_type, user_pass='', owner_pass=''):
        self.pattern = pattern
        self.replacement = replacement
        self.priority = priority
        self.rule_type = rule_type
        self.user_pass = user_pass
        self.owner_pass = owner_pass