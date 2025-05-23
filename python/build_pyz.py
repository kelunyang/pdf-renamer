#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import zipapp
import shutil
import argparse
from pathlib import Path

def build_pyz(output_name="pdfRenamer", include_paddleocr=False):
    """将程序打包成pyz格式
    
    参数:
        output_name (str): 输出文件名（不含扩展名）
        include_paddleocr (bool): 是否包含PaddleOCR（会显著增加包大小）
    """
    # 创建临时目录
    build_dir = Path("build")
    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir.mkdir()
    
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
    
    # 复制所有Python文件到build目录
    for file in py_files:
        if Path(file).exists():
            shutil.copy(file, build_dir)
    
    # 创建__main__.py文件作为入口点，直接复制main.py的内容
    if Path("main.py").exists():
        with open("main.py", "r", encoding="utf-8") as src_file:
            main_content = src_file.read()
        
        with open(build_dir / "__main__.py", "w", encoding="utf-8") as f:
            f.write(main_content)
    else:
        print("错误：找不到main.py文件")
        shutil.rmtree(build_dir)  # 清理临时目录
        return None
    
    # 创建requirements.txt文件
    with open(build_dir / "requirements.txt", "w", encoding="utf-8") as f:
        requirements = [
            "PyPDF2",
            "PyMuPDF",  # fitz
            "pikepdf",
            "questionary",
            "tqdm"
        ]
        
        if include_paddleocr:
            requirements.extend(["paddleocr", "paddlepaddle"])
        
        f.write("\n".join(requirements))
    
    # 使用zipapp创建可执行的pyz文件
    output_file = f"{output_name}.pyz"
    zipapp.create_archive(
        build_dir,
        output_file
    )
    
    print(f"打包完成: {output_file}")
    print(f"文件大小: {os.path.getsize(output_file) / 1024 / 1024:.2f} MB")
    
    # 清理临时目录
    shutil.rmtree(build_dir)
    
    return output_file

def main():
    parser = argparse.ArgumentParser(description="将PDF重命名工具打包为pyz格式")
    parser.add_argument(
        "-o", "--output", 
        default="pdfRenamer", 
        help="输出文件名（不含扩展名）"
    )
    parser.add_argument(
        "--include-paddleocr", 
        action="store_true", 
        help="包含PaddleOCR（会显著增加包大小）"
    )
    
    args = parser.parse_args()
    build_pyz(args.output, args.include_paddleocr)

if __name__ == "__main__":
    main()