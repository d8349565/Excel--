# -*- coding: utf-8 -*-
"""
用户操作日志记录器
记录用户的详细操作记录，包括文件上传、删除、合并等操作
"""

import os
import logging
import json
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Optional
from functools import wraps
from flask import session, request

class UserLogger:
    """用户操作日志记录器"""
    
    def __init__(self, app_config: dict):
        self.logs_folder = app_config.get('LOGS_FOLDER', 'logs')
        self.user_log_file = os.path.join(self.logs_folder, 'user_operations.log')
        
        # 设置北京时区
        self.beijing_tz = timezone(timedelta(hours=8))
        
        # 确保日志文件夹存在
        if not os.path.exists(self.logs_folder):
            os.makedirs(self.logs_folder)
        
        # 配置用户操作日志器
        self.logger = logging.getLogger('user_operations')
        self.logger.setLevel(logging.INFO)
        
        # 避免重复添加handler
        if not self.logger.handlers:
            handler = logging.FileHandler(self.user_log_file, encoding='utf-8')
            formatter = logging.Formatter(
                '%(asctime)s | %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
    
    def log_operation(self, operation: str, details: dict = None, session_id: str = None):
        """记录用户操作"""
        try:
            # 获取会话ID
            if session_id is None:
                session_id = session.get('session_id', 'unknown')
            
            # 获取客户端IP
            client_ip = request.remote_addr if request else 'unknown'
            user_agent = request.headers.get('User-Agent', 'unknown') if request else 'unknown'
            
            # 构建日志记录
            log_entry = {
                'session_id': session_id,
                'operation': operation,
                'client_ip': client_ip,
                'user_agent': user_agent[:100],  # 限制长度
                'details': details or {},
                'timestamp': datetime.now(self.beijing_tz).isoformat()
            }
            
            # 记录到日志文件
            self.logger.info(json.dumps(log_entry, ensure_ascii=False))
            
        except Exception as e:
            # 日志记录失败不应该影响主要功能
            print(f"用户操作日志记录失败: {e}")
    
    def log_file_upload(self, filename: str, file_size: int, file_id: str, session_id: str = None):
        """记录文件上传"""
        details = {
            'filename': filename,
            'file_size': file_size,
            'file_id': file_id,
            'file_size_mb': round(file_size / (1024 * 1024), 2)
        }
        self.log_operation('file_upload', details, session_id)
    
    def log_file_delete(self, filename: str, file_id: str, session_id: str = None):
        """记录文件删除"""
        details = {
            'filename': filename,
            'file_id': file_id
        }
        self.log_operation('file_delete', details, session_id)
    
    def log_file_preview(self, filename: str, file_id: str, sheet_name: str = None, session_id: str = None):
        """记录文件预览"""
        details = {
            'filename': filename,
            'file_id': file_id,
            'sheet_name': sheet_name
        }
        self.log_operation('file_preview', details, session_id)
    
    def log_clear_all_files(self, files_count: int, session_id: str = None):
        """记录清空所有文件"""
        details = {
            'files_count': files_count
        }
        self.log_operation('clear_all_files', details, session_id)
    
    def log_merge_task_submit(self, file_count: int, task_id: str, export_format: str = None, session_id: str = None):
        """记录合并任务提交"""
        details = {
            'file_count': file_count,
            'task_id': task_id,
            'export_format': export_format
        }
        self.log_operation('merge_task_submit', details, session_id)
    
    def log_merge_task_complete(self, task_id: str, result_filename: str, processing_time: float = None, session_id: str = None):
        """记录合并任务完成"""
        details = {
            'task_id': task_id,
            'result_filename': result_filename,
            'processing_time_seconds': processing_time
        }
        self.log_operation('merge_task_complete', details, session_id)
    
    def log_file_download(self, filename: str, session_id: str = None):
        """记录文件下载"""
        details = {
            'filename': filename
        }
        self.log_operation('file_download', details, session_id)
    
    def log_login(self, session_id: str = None):
        """记录用户登录"""
        self.log_operation('user_login', {}, session_id)
    
    def log_logout(self, session_id: str = None):
        """记录用户登出"""
        self.log_operation('user_logout', {}, session_id)
    
    def get_user_logs(self, session_id: str = None, limit: int = 100, operation_filter: str = None) -> List[Dict]:
        """获取用户操作日志"""
        logs = []
        
        if not os.path.exists(self.user_log_file):
            return logs
        
        try:
            with open(self.user_log_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # 倒序读取（最新的在前）
            for line in reversed(lines[-limit:]):
                try:
                    # 解析日志行
                    if ' | ' in line:
                        timestamp_str, json_str = line.strip().split(' | ', 1)
                        log_data = json.loads(json_str)
                        
                        # 过滤条件
                        if session_id and log_data.get('session_id') != session_id:
                            continue
                        
                        if operation_filter and log_data.get('operation') != operation_filter:
                            continue
                        
                        # 添加格式化的时间戳
                        log_data['formatted_time'] = timestamp_str
                        logs.append(log_data)
                        
                        if len(logs) >= limit:
                            break
                            
                except (json.JSONDecodeError, ValueError):
                    # 跳过无效的日志行
                    continue
            
        except Exception as e:
            print(f"读取用户日志失败: {e}")
        
        return logs
    
    def get_operation_stats(self, session_id: str = None, days: int = 7) -> Dict:
        """获取操作统计信息"""
        from collections import defaultdict
        
        stats = {
            'total_operations': 0,
            'operations_by_type': defaultdict(int),
            'files_uploaded': 0,
            'files_deleted': 0,
            'merge_tasks': 0,
            'last_activity': None
        }
        
        # 计算时间范围
        from datetime import timedelta
        cutoff_time = datetime.now() - timedelta(days=days)
        
        logs = self.get_user_logs(session_id=session_id, limit=1000)
        
        for log in logs:
            try:
                # 解析时间戳
                log_time = datetime.fromisoformat(log.get('timestamp', ''))
                
                if log_time < cutoff_time:
                    continue
                
                operation = log.get('operation', '')
                stats['total_operations'] += 1
                stats['operations_by_type'][operation] += 1
                
                # 统计特定操作
                if operation == 'file_upload':
                    stats['files_uploaded'] += 1
                elif operation == 'file_delete':
                    stats['files_deleted'] += 1
                elif operation == 'merge_task_submit':
                    stats['merge_tasks'] += 1
                
                # 记录最后活动时间
                if stats['last_activity'] is None or log_time > stats['last_activity']:
                    stats['last_activity'] = log_time
                    
            except (ValueError, TypeError):
                continue
        
        # 转换为普通字典
        stats['operations_by_type'] = dict(stats['operations_by_type'])
        
        return stats

def log_user_operation(operation: str, **details):
    """装饰器：自动记录用户操作"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                # 执行原函数
                result = func(*args, **kwargs)
                
                # 记录操作（需要确保在应用上下文中）
                from flask import current_app
                if hasattr(current_app, 'user_logger'):
                    current_app.user_logger.log_operation(operation, details)
                
                return result
            except Exception as e:
                # 即使记录失败也不影响原函数执行
                return func(*args, **kwargs)
        return wrapper
    return decorator