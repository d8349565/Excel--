#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
服务器启动文件
用于生产环境部署
注意安装gunicorn后使用
gunicorn -w 4 -b 0.0.0.0:5000 serve:app --chdir excel_tool
"""
# pkill -f gunicorn
# /home/devbox/project/bin/gunicorn -w 4 -b 0.0.0.0:5000 serve:app --chdir excel_tool

import os
import sys

# 添加excel_tool目录到Python路径
excel_tool_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, excel_tool_dir)

# 设置工作目录
os.chdir(excel_tool_dir)

# 确保任务存储目录存在
task_storage_dir = os.path.join(excel_tool_dir, 'task_storage')
os.makedirs(task_storage_dir, exist_ok=True)

from app import create_app

# 创建应用实例
app = create_app('production')

if __name__ == '__main__':
    # 开发环境启动
    app.run(host='0.0.0.0', port=5000, debug=False)