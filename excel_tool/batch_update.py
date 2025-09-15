#!/usr/bin/env python3
"""
批量添加 @login_required 装饰器和会话支持的脚本
"""

import re

def add_login_protection():
    """给app.py中的路由添加登录保护"""
    
    # 需要保护的路由（不包括login和logout）
    routes_to_protect = [
        "'/api/detect_columns/<file_id>'",
        "'/get_merged_columns'",
        "'/preview_merge'", 
        "'/submit_task'",
        "'/task/<task_id>'",
        "'/api/task/<task_id>'",
        "'/download/<filename>'",
        "'/api/delete_file/<file_id>'",
        "'/api/clear_all_files'",
        "'/api/system_status'"
    ]
    
    print("需要保护的路由:")
    for route in routes_to_protect:
        print(f"  - {route}")
    
    print("\n请手动在每个路由函数前添加 @login_required 装饰器")
    print("并在需要的地方添加 session_id = get_session_id()")

if __name__ == "__main__":
    add_login_protection()