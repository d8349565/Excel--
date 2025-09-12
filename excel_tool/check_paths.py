#!/usr/bin/env python3
"""
路径检查脚本 - 验证文件夹路径配置
"""

import os
import sys
from config import config

def check_paths():
    """检查路径配置"""
    print("=" * 60)
    print("Excel汇总工具 - 路径配置检查")
    print("=" * 60)
    
    # 获取当前工作目录
    current_dir = os.getcwd()
    print(f"当前工作目录: {current_dir}")
    
    # 获取脚本所在目录
    script_dir = os.path.abspath(os.path.dirname(__file__))
    print(f"脚本所在目录: {script_dir}")
    
    # 检查配置
    config_obj = config['default']
    
    print(f"\n配置路径:")
    print(f"UPLOAD_FOLDER: {config_obj.UPLOAD_FOLDER}")
    print(f"RESULTS_FOLDER: {config_obj.RESULTS_FOLDER}")
    print(f"LOGS_FOLDER: {config_obj.LOGS_FOLDER}")
    
    print(f"\n绝对路径:")
    print(f"上传文件夹: {os.path.abspath(config_obj.UPLOAD_FOLDER)}")
    print(f"结果文件夹: {os.path.abspath(config_obj.RESULTS_FOLDER)}")
    print(f"日志文件夹: {os.path.abspath(config_obj.LOGS_FOLDER)}")
    
    print(f"\n文件夹存在性检查:")
    for name, path in [
        ("上传文件夹", config_obj.UPLOAD_FOLDER),
        ("结果文件夹", config_obj.RESULTS_FOLDER),
        ("日志文件夹", config_obj.LOGS_FOLDER)
    ]:
        abs_path = os.path.abspath(path)
        exists = os.path.exists(abs_path)
        print(f"{name}: {'✓ 存在' if exists else '✗ 不存在'} - {abs_path}")
        
        if not exists:
            try:
                os.makedirs(abs_path, exist_ok=True)
                print(f"  → 已创建文件夹: {abs_path}")
            except Exception as e:
                print(f"  → 创建失败: {e}")
    
    print("\n" + "=" * 60)

if __name__ == "__main__":
    check_paths()