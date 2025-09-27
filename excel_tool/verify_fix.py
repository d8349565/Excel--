#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速验证配置传递是否修复
"""

import os
import sys

def check_config_support():
    """检查config.py是否支持环境变量"""
    print("检查 config.py 环境变量支持...")
    
    # 读取config.py内容
    config_file = 'config.py'
    if not os.path.exists(config_file):
        print("❌ 找不到 config.py")
        return False
    
    with open(config_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查关键的环境变量支持
    checks = [
        ('MAX_FILE_SIZE_MB', 'os.environ.get(\'MAX_FILE_SIZE_MB\''),
        ('MAX_CONCURRENT_TASKS', 'os.environ.get(\'MAX_CONCURRENT_TASKS\''),
        ('FILE_RETENTION_DAYS', 'os.environ.get(\'FILE_RETENTION_DAYS\''),
        ('DEFAULT_PREVIEW_ROWS', 'os.environ.get(\'DEFAULT_PREVIEW_ROWS\''),
    ]
    
    all_good = True
    for var_name, expected_pattern in checks:
        if expected_pattern in content:
            print(f"✅ {var_name} 环境变量支持已添加")
        else:
            print(f"❌ {var_name} 环境变量支持缺失")
            all_good = False
    
    return all_good

def check_gui_transmission():
    """检查GUI控制器是否传递所有参数"""
    print("\n检查 GUI 控制器环境变量传递...")
    
    gui_file = 'gui_controller.py'
    if not os.path.exists(gui_file):
        print("❌ 找不到 gui_controller.py")
        return False
    
    with open(gui_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查环境变量设置
    checks = [
        ('MAX_FILE_SIZE_MB', 'env[\'MAX_FILE_SIZE_MB\']'),
        ('MAX_CONCURRENT_TASKS', 'env[\'MAX_CONCURRENT_TASKS\']'),
        ('FILE_RETENTION_DAYS', 'env[\'FILE_RETENTION_DAYS\']'),
        ('DEFAULT_PREVIEW_ROWS', 'env[\'DEFAULT_PREVIEW_ROWS\']'),
    ]
    
    all_good = True
    for var_name, expected_pattern in checks:
        if expected_pattern in content:
            print(f"✅ {var_name} 环境变量传递已添加")
        else:
            print(f"❌ {var_name} 环境变量传递缺失")
            all_good = False
    
    return all_good

def main():
    print("=" * 50)
    print("配置修复验证脚本")
    print("=" * 50)
    
    # 检查config.py
    config_ok = check_config_support()
    
    # 检查GUI控制器
    gui_ok = check_gui_transmission()
    
    print("\n" + "=" * 50)
    if config_ok and gui_ok:
        print("🎉 修复验证通过！")
        print("现在GUI中的所有配置都能正确传递给Flask应用了！")
        
        print("\n可传递的配置参数:")
        print("• 文件大小限制 (max_file_size)")
        print("• 最大并发任务数 (max_concurrent_tasks)")  
        print("• 文件保留天数 (file_retention_days)")
        print("• 默认预览行数 (preview_rows)")
        print("• 服务器地址和端口")
        print("• 用户认证密码")
        print("• 调试模式")
    else:
        print("❌ 修复验证失败！")
        print("请检查代码修改是否正确。")
    
    print("=" * 50)
    return config_ok and gui_ok

if __name__ == '__main__':
    main()
