import warnings
import logging
import os
import sys
from log_utils import log_message

# 創建一個自定義的警告過濾器
class PaddleWarningFilter(warnings.catch_warnings):
    """自定義警告過濾器，用於捕獲PaddleOCR的警告訊息並重定向到日誌系統"""
    
    def __enter__(self):
        # 調用父類的__enter__方法
        self._record = super().__enter__()
        # 設置警告過濾器
        warnings.filterwarnings('always', category=UserWarning)
        # 返回記錄對象
        return self._record

# 創建一個自定義的警告處理函數
def paddle_warning_handler(message, category, filename, lineno, file=None, line=None):
    """自定義警告處理函數，將警告訊息記錄到日誌系統"""
    # 將警告訊息記錄到日誌系統
    warning_message = f"{category.__name__}: {message}"
    log_message(warning_message, level='警告')
    # 返回None表示不顯示警告訊息
    return None

# 創建一個自定義的日誌處理器
class PaddleLogHandler(logging.Handler):
    """自定義日誌處理器，將PaddleOCR的日誌訊息重定向到我們的日誌系統"""
    
    def emit(self, record):
        """處理日誌記錄"""
        # 將日誌訊息記錄到我們的日誌系統
        level_map = {
            logging.DEBUG: '調試',
            logging.INFO: '信息',
            logging.WARNING: '警告',
            logging.ERROR: '错误',
            logging.CRITICAL: '嚴重'
        }
        level = level_map.get(record.levelno, '信息')
        log_message(record.getMessage(), level=level)

# 設置PaddleOCR的日誌級別
def setup_paddle_logging():
    """設置PaddleOCR的日誌級別和處理器"""
    # 設置警告處理函數
    warnings.showwarning = paddle_warning_handler
    
    # 設置Paddle相關模塊的日誌級別
    paddle_loggers = ['paddle', 'ppocr', 'paddleocr']
    for logger_name in paddle_loggers:
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.WARNING)  # 只顯示警告及以上級別的日誌
        
        # 移除所有現有的處理器
        for handler in logger.handlers[:]:  # 使用切片創建副本，避免在迭代時修改
            logger.removeHandler(handler)
        
        # 添加我們的自定義處理器
        handler = PaddleLogHandler()
        logger.addHandler(handler)
        
        # 設置不向上傳播日誌
        logger.propagate = False
    
    # 記錄設置完成的訊息
    log_message("已設置PaddleOCR的日誌重定向", level='信息')

# 初始化PaddleOCR時調用此函數
def init_paddleocr(use_angle_cls=True, lang="ch"):
    """初始化PaddleOCR並設置日誌重定向"""
    # 首先設置日誌重定向
    setup_paddle_logging()
    
    try:
        # 導入PaddleOCR
        from paddleocr import PaddleOCR
        
        # 使用自定義的警告過濾器
        with PaddleWarningFilter() as w:
            # 初始化PaddleOCR
            ocr = PaddleOCR(use_angle_cls=use_angle_cls, lang=lang)
            
            # 檢查OCR實例是否有效
            if ocr is None:
                log_message("PaddleOCR初始化返回了None", level='错误')
                return None
                
            # 檢查OCR實例是否有必要的方法
            if not hasattr(ocr, 'ocr'):
                log_message("PaddleOCR實例缺少ocr方法", level='错误')
                return None
            
            # 如果有警告，記錄到日誌系統
            if w is not None:
                try:
                    for warning in w:
                        log_message(str(warning.message), level='警告')
                except Exception as warn_err:
                    log_message(f"處理PaddleOCR警告時出錯: {warn_err}", level='警告')
        
        return ocr
    except ImportError:
        log_message("無法導入PaddleOCR，請確保已安裝paddleocr和paddlepaddle", level='错误')
        return None
    except Exception as e:
        log_message(f"初始化PaddleOCR時出錯: {e}", level='错误')
        return None