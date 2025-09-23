import threading
import queue
import uuid
import time
import logging
from datetime import datetime
from typing import Dict, Any, Optional, Callable
from enum import Enum
import json
import os

class TaskStatus(Enum):
    """任务状态枚举"""
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    TIMEOUT = "timeout"

class Task:
    """任务对象"""
    
    def __init__(self, task_id: str, task_type: str, parameters: Dict[str, Any]):
        self.task_id = task_id
        self.task_type = task_type
        self.parameters = parameters
        self.status = TaskStatus.PENDING
        self.created_at = datetime.now()
        self.started_at = None
        self.completed_at = None
        self.progress = 0
        self.result = None
        self.error_message = None
        self.logs = []
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'task_id': self.task_id,
            'task_type': self.task_type,
            'parameters': self.parameters,
            'status': self.status.value,
            'created_at': self.created_at.isoformat(),
            'started_at': self.started_at.isoformat() if self.started_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'progress': self.progress,
            'result': self.result,
            'error_message': self.error_message,
            'logs': self.logs
        }

class TaskManager:
    """任务管理器，处理后台任务执行"""
    
    def __init__(self, max_workers: int = 1, task_timeout: int = 3600):
        self.max_workers = max_workers
        self.task_timeout = task_timeout
        self.tasks: Dict[str, Task] = {}
        self.task_queue = queue.Queue()
        self.workers = []
        self.running = False
        self.lock = threading.Lock()
        
        # 注册的任务处理器
        self.task_handlers: Dict[str, Callable] = {}
    
    def start(self):
        """启动任务管理器"""
        if self.running:
            return
        
        self.running = True
        
        # 启动工作线程
        for i in range(self.max_workers):
            worker = threading.Thread(target=self._worker_loop, name=f"TaskWorker-{i}")
            worker.daemon = True
            worker.start()
            self.workers.append(worker)
        
        logging.info(f"任务管理器已启动，{self.max_workers} 个工作线程")
    
    def stop(self):
        """停止任务管理器"""
        self.running = False
        
        # 向队列添加停止信号
        for _ in range(self.max_workers):
            self.task_queue.put(None)
        
        # 等待工作线程结束
        for worker in self.workers:
            worker.join(timeout=5)
        
        logging.info("任务管理器已停止")
    
    def register_handler(self, task_type: str, handler: Callable):
        """
        注册任务处理器
        
        Args:
            task_type: 任务类型
            handler: 处理函数，接收(task, progress_callback)参数
        """
        self.task_handlers[task_type] = handler
        logging.info(f"已注册任务处理器: {task_type}")
    
    def submit_task(self, task_type: str, parameters: Dict[str, Any]) -> str:
        """
        提交任务
        
        Args:
            task_type: 任务类型
            parameters: 任务参数
            
        Returns:
            任务ID
        """
        if task_type not in self.task_handlers:
            raise ValueError(f"未注册的任务类型: {task_type}")
        
        # 创建任务
        task_id = str(uuid.uuid4())
        task = Task(task_id, task_type, parameters)
        
        # 保存任务
        with self.lock:
            self.tasks[task_id] = task
        
        # 添加到队列
        self.task_queue.put(task_id)
        
        logging.info(f"任务已提交: {task_id} ({task_type})")
        return task_id
    
    def get_task(self, task_id: str) -> Optional[Task]:
        """获取任务信息"""
        with self.lock:
            return self.tasks.get(task_id)
    
    def get_task_status(self, task_id: str) -> Optional[Dict[str, Any]]:
        """获取任务状态"""
        task = self.get_task(task_id)
        return task.to_dict() if task else None
    
    def get_all_tasks(self) -> Dict[str, Dict[str, Any]]:
        """获取所有任务状态"""
        with self.lock:
            return {task_id: task.to_dict() for task_id, task in self.tasks.items()}
    
    def cancel_task(self, task_id: str) -> bool:
        """取消任务（仅对未开始的任务有效）"""
        with self.lock:
            task = self.tasks.get(task_id)
            if task and task.status == TaskStatus.PENDING:
                task.status = TaskStatus.FAILED
                task.error_message = "任务已取消"
                task.completed_at = datetime.now()
                return True
        return False
    
    def _worker_loop(self):
        """工作线程主循环"""
        while self.running:
            try:
                # 从队列获取任务
                task_id = self.task_queue.get(timeout=1)
                
                if task_id is None:  # 停止信号
                    break
                
                # 执行任务
                self._execute_task(task_id)
                
            except queue.Empty:
                continue
            except Exception as e:
                logging.error(f"工作线程出错: {e}")
    
    def _execute_task(self, task_id: str):
        """执行单个任务"""
        task = self.get_task(task_id)
        if not task:
            return
        
        try:
            # 更新任务状态
            with self.lock:
                task.status = TaskStatus.RUNNING
                task.started_at = datetime.now()
                task.progress = 0
            
            logging.info(f"开始执行任务: {task_id}")
            
            # 创建进度回调函数
            def progress_callback(progress: int, message: str = ""):
                with self.lock:
                    task.progress = max(0, min(100, progress))
                    if message:
                        task.logs.append({
                            'timestamp': datetime.now().isoformat(),
                            'message': message
                        })
            
            # 获取任务处理器
            handler = self.task_handlers.get(task.task_type)
            if not handler:
                raise ValueError(f"未找到任务处理器: {task.task_type}")
            
            # 设置超时
            start_time = time.time()
            
            # 在单独线程中执行任务处理器
            result_queue = queue.Queue()
            
            def run_handler():
                try:
                    result = handler(task, progress_callback)
                    result_queue.put(('success', result))
                except Exception as e:
                    result_queue.put(('error', str(e)))
            
            handler_thread = threading.Thread(target=run_handler)
            handler_thread.start()
            
            # 等待任务完成或超时
            while handler_thread.is_alive():
                if time.time() - start_time > self.task_timeout:
                    # 任务超时
                    with self.lock:
                        task.status = TaskStatus.TIMEOUT
                        task.error_message = f"任务超时（{self.task_timeout}秒）"
                        task.completed_at = datetime.now()
                    logging.warning(f"任务超时: {task_id}")
                    return
                
                time.sleep(1)
            
            # 获取结果
            try:
                result_type, result_data = result_queue.get_nowait()
                
                if result_type == 'success':
                    with self.lock:
                        task.status = TaskStatus.COMPLETED
                        task.result = result_data
                        task.progress = 100
                        task.completed_at = datetime.now()
                    logging.info(f"任务完成: {task_id}")
                else:
                    with self.lock:
                        task.status = TaskStatus.FAILED
                        task.error_message = result_data
                        task.completed_at = datetime.now()
                    logging.error(f"任务失败: {task_id} - {result_data}")
                    
            except queue.Empty:
                with self.lock:
                    task.status = TaskStatus.FAILED
                    task.error_message = "任务处理器未返回结果"
                    task.completed_at = datetime.now()
        
        except Exception as e:
            with self.lock:
                task.status = TaskStatus.FAILED
                task.error_message = str(e)
                task.completed_at = datetime.now()
            logging.error(f"执行任务时出错: {task_id} - {e}")
    
    def cleanup_old_tasks(self, max_age_hours: int = 24):
        """清理旧任务"""
        cutoff_time = datetime.now().timestamp() - (max_age_hours * 3600)
        
        with self.lock:
            to_remove = []
            for task_id, task in self.tasks.items():
                if task.created_at.timestamp() < cutoff_time:
                    to_remove.append(task_id)
            
            for task_id in to_remove:
                del self.tasks[task_id]
        
        if to_remove:
            logging.info(f"清理了 {len(to_remove)} 个旧任务")
    
    def get_queue_size(self) -> int:
        """获取队列大小"""
        return self.task_queue.qsize()
    
    def get_running_tasks(self) -> int:
        """获取正在运行的任务数量"""
        with self.lock:
            return sum(1 for task in self.tasks.values() 
                      if task.status == TaskStatus.RUNNING)

# 全局任务管理器实例
task_manager = TaskManager()

def create_merge_task_handler(file_manager, data_processor):
    """
    创建数据合并任务处理器
    
    Args:
        file_manager: 文件管理器实例
        data_processor: 数据处理器实例
        
    Returns:
        任务处理器函数
    """
    def handle_merge_task(task: Task, progress_callback: Callable):
        """处理数据合并任务"""
        try:
            parameters = task.parameters
            file_configs = parameters.get('file_configs', [])
            cleaning_options = parameters.get('cleaning_options', {})
            export_options = parameters.get('export_options', {})
            session_id = parameters.get('session_id')  # 获取会话ID
            
            if not file_configs:
                raise ValueError("没有指定要处理的文件")
            
            progress_callback(10, "开始读取文件...")
            
            # 读取所有文件数据
            dataframes = []
            total_files = len(file_configs)
            
            for i, config in enumerate(file_configs):
                file_id = config['file_id']
                sheet_name = config.get('sheet_name')
                header_row = config.get('header_row', 0)
                source_name = config.get('source_name', f"文件{i+1}")
                
                try:
                    df = file_manager.read_full_file(file_id, sheet_name, header_row, session_id)
                    dataframes.append((df, source_name))
                    
                    progress_callback(
                        10 + (i + 1) * 30 // total_files,
                        f"已读取文件: {source_name}"
                    )
                except Exception as e:
                    raise ValueError(f"读取文件失败 {source_name}: {str(e)}")
            
            progress_callback(40, "开始数据处理...")
            
            # 处理数据，传递file_manager和session_id以支持固定单元格功能
            result_df = data_processor.process_data(dataframes, cleaning_options, file_manager, session_id)
            
            progress_callback(80, "开始导出文件...")
            
            # 导出结果
            export_format = export_options.get('format', 'xlsx')
            export_filename = export_options.get('filename', 'merged_data')
            
            result_path = file_manager.save_result_file(
                result_df, export_filename, export_format, session_id
            )
            
            # 确保结果文件存在
            if not os.path.exists(result_path):
                raise Exception(f"结果文件保存失败: {result_path}")
            
            progress_callback(95, "生成处理报告...")
            
            # 生成处理摘要
            summary = data_processor.get_processing_summary()
            
            progress_callback(100, "任务完成")
            
            result_filename = os.path.basename(result_path)
            
            return {
                'result_file': result_filename,
                'result_path': result_path,
                'summary': summary,
                'export_format': export_format,
                'total_rows': len(result_df),
                'total_columns': len(result_df.columns)
            }
            
        except Exception as e:
            logging.error(f"处理合并任务时出错: {e}")
            raise
    
    return handle_merge_task