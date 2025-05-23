#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import argparse
import subprocess
from pathlib import Path

def build_exe(output_name="pdfRenamer", include_paddleocr=False, one_file=True, console=True):
    """
    將程序打包成exe格式
    
    參數:
        output_name (str): 輸出文件名（不含擴展名）
        include_paddleocr (bool): 是否包含PaddleOCR（會顯著增加包大小）
        one_file (bool): 是否打包成單個文件
        console (bool): 是否顯示控制台窗口
    """
    # 檢查是否安裝了PyInstaller
    try:
        import PyInstaller
        print(f"使用PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("未安裝PyInstaller，正在嘗試安裝...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller安裝成功")
        except subprocess.CalledProcessError as e:
            print(f"安裝PyInstaller失敗: {e}")
            return None
    
    # 創建臨時目錄
    build_dir = Path("build")
    dist_dir = Path("dist")
    if build_dir.exists():
        shutil.rmtree(build_dir)
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    
    # 要包含的Python文件
    py_files = [
        "main.py",
        "db_utils.py",
        "log_utils.py",
        "file_utils.py",
        "pdf_utils.py",
        "input_utils.py",
        "rule_utils.py",
        "worker_utils.py",
        "ui_utils.py",
        "__init__.py"
    ]
    
    # 如果包含PaddleOCR，添加paddle_utils.py
    if include_paddleocr and Path("paddle_utils.py").exists():
        py_files.append("paddle_utils.py")
    
    # 檢查所有文件是否存在
    missing_files = [file for file in py_files if not Path(file).exists()]
    if missing_files:
        print(f"警告：以下文件不存在: {', '.join(missing_files)}")
    
    # 構建PyInstaller命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", output_name,
        "--clean"
    ]
    
    # 添加選項
    if one_file:
        cmd.append("--onefile")
    else:
        cmd.append("--onedir")
    
    if not console:
        cmd.append("--noconsole")
    
    # 添加數據文件
    # 如果有其他數據文件需要包含，可以在這裡添加
    # cmd.extend(["--add-data", "path/to/data;data"])
    
    # 添加隱藏導入
    hidden_imports = [
        "questionary",
        "tqdm",
        "PyPDF2",
        "pikepdf",
        "fitz",  # PyMuPDF
        "sqlite3"
    ]
    
    if include_paddleocr:
        hidden_imports.extend(["paddleocr", "paddlepaddle"])
    
    for imp in hidden_imports:
        cmd.extend(["--hidden-import", imp])
    
    # 添加主腳本
    cmd.append("main.py")
    
    # 執行PyInstaller命令
    print("正在執行PyInstaller打包...")
    print(f"命令: {' '.join(cmd)}")
    try:
        subprocess.check_call(cmd)
        print("打包完成!")
        
        # 獲取輸出文件路徑
        if one_file:
            output_file = dist_dir / f"{output_name}.exe"
        else:
            output_file = dist_dir / output_name / f"{output_name}.exe"
        
        if output_file.exists():
            print(f"輸出文件: {output_file}")
            print(f"文件大小: {output_file.stat().st_size / 1024 / 1024:.2f} MB")
            return output_file
        else:
            print(f"錯誤：找不到輸出文件 {output_file}")
            return None
    except subprocess.CalledProcessError as e:
        print(f"打包失敗: {e}")
        return None

def main():
    parser = argparse.ArgumentParser(description="將PDF重命名工具打包為exe格式")
    parser.add_argument(
        "-o", "--output", 
        default="pdfRenamer", 
        help="輸出文件名（不含擴展名）"
    )
    parser.add_argument(
        "--include-paddleocr", 
        action="store_true", 
        help="包含PaddleOCR（會顯著增加包大小）"
    )
    parser.add_argument(
        "--onedir", 
        action="store_false", 
        dest="one_file",
        help="打包為目錄而非單個文件"
    )
    parser.add_argument(
        "--noconsole", 
        action="store_false", 
        dest="console",
        help="不顯示控制台窗口"
    )
    
    args = parser.parse_args()
    build_exe(args.output, args.include_paddleocr, args.one_file, args.console)

if __name__ == "__main__":
    main()