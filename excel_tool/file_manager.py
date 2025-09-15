import os
import uuid
import shutil
import time
import pandas as pd
import chardet
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import logging
from typing import List, Dict, Optional, Tuple
import json

class FileManager:
    """文件管理器，处理文件上传、预览、临时文件管理等功能"""
    
    def __init__(self, app=None):
        self.app = app
        self.upload_folder = None
        self.results_folder = None
        self.allowed_extensions = None
        self.max_file_size = None
        
        if app:
            self.init_app(app)
    
    def init_app(self, app):
        """初始化应用配置"""
        self.app = app
        self.upload_folder = app.config['UPLOAD_FOLDER']
        self.results_folder = app.config['RESULTS_FOLDER']
        self.allowed_extensions = app.config['ALLOWED_EXTENSIONS']
        self.max_file_size = app.config['MAX_FILE_SIZE']
        
        # 记录路径信息
        logging.info(f"FileManager初始化 - 上传文件夹: {self.upload_folder}")
        logging.info(f"FileManager初始化 - 结果文件夹: {self.results_folder}")
        
        # 确保目录存在
        os.makedirs(self.upload_folder, exist_ok=True)
        os.makedirs(self.results_folder, exist_ok=True)
        
        # 验证目录是否成功创建
        if os.path.exists(self.upload_folder):
            logging.info(f"上传文件夹已就绪: {os.path.abspath(self.upload_folder)}")
        else:
            logging.error(f"上传文件夹创建失败: {self.upload_folder}")
            
        if os.path.exists(self.results_folder):
            logging.info(f"结果文件夹已就绪: {os.path.abspath(self.results_folder)}")
        else:
            logging.error(f"结果文件夹创建失败: {self.results_folder}")
    
    def get_user_upload_folder(self, session_id):
        """获取用户专用上传文件夹"""
        user_folder = os.path.join(self.upload_folder, session_id)
        os.makedirs(user_folder, exist_ok=True)
        return user_folder
    
    def get_user_results_folder(self, session_id):
        """获取用户专用结果文件夹"""
        user_folder = os.path.join(self.results_folder, session_id)
        os.makedirs(user_folder, exist_ok=True)
        return user_folder
    
    def allowed_file(self, filename: str) -> bool:
        """检查文件扩展名是否被允许"""
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in self.allowed_extensions
    
    def get_file_size(self, file) -> int:
        """获取文件大小"""
        file.seek(0, 2)  # 移动到文件末尾
        size = file.tell()
        file.seek(0)     # 回到文件开头
        return size
    
    def save_uploaded_file(self, file, original_filename: str, session_id: str = None) -> Dict:
        """
        保存上传的文件
        
        Args:
            file: 上传的文件对象
            original_filename: 原始文件名
            session_id: 用户会话ID，用于文件隔离
            
        Returns:
            包含文件信息的字典
        """
        try:
            # 检查文件扩展名
            if not self.allowed_file(original_filename):
                raise ValueError(f"不支持的文件类型: {original_filename}")
            
            # 检查文件大小
            file_size = self.get_file_size(file)
            if file_size > self.max_file_size:
                raise ValueError(f"文件过大: {file_size / (1024*1024):.1f}MB，最大允许 {self.max_file_size / (1024*1024):.1f}MB")
            
            # 生成唯一文件ID和安全文件名
            file_id = str(uuid.uuid4())
            safe_filename = secure_filename(original_filename)
            extension = safe_filename.rsplit('.', 1)[1].lower()
            saved_filename = f"{file_id}.{extension}"
            
            # 根据是否有session_id决定保存位置
            if session_id:
                user_upload_folder = self.get_user_upload_folder(session_id)
                file_path = os.path.join(user_upload_folder, saved_filename)
            else:
                file_path = os.path.join(self.upload_folder, saved_filename)
            
            # 保存文件
            file.save(file_path)
            
            # 获取文件信息
            file_info = {
                'file_id': file_id,
                'original_filename': original_filename,
                'saved_filename': saved_filename,
                'file_path': file_path,
                'file_size': file_size,
                'extension': extension,
                'upload_time': datetime.now().isoformat(),
                'session_id': session_id,
                'sheets': []
            }
            
            # 如果是Excel文件，获取sheet列表
            if extension in ['xlsx', 'xls']:
                try:
                    sheets = pd.ExcelFile(file_path).sheet_names
                    file_info['sheets'] = sheets
                except Exception as e:
                    logging.warning(f"无法读取Excel文件的sheet信息: {e}")
                    file_info['sheets'] = ['Sheet1']  # 默认sheet名
            else:
                file_info['sheets'] = ['CSV']  # CSV文件只有一个"sheet"
            
            # 保存文件元信息
            self._save_file_metadata(file_info)
            
            return file_info
            
        except Exception as e:
            logging.error(f"保存文件时出错: {e}")
            raise
    
    def get_file_info(self, file_id: str, session_id: str = None) -> Optional[Dict]:
        """获取文件信息"""
        if session_id:
            user_folder = self.get_user_upload_folder(session_id)
            metadata_path = os.path.join(user_folder, f"{file_id}_metadata.json")
        else:
            metadata_path = os.path.join(self.upload_folder, f"{file_id}_metadata.json")
            
        if os.path.exists(metadata_path):
            try:
                with open(metadata_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"读取文件元信息失败: {e}")
        return None
    
    def _save_file_metadata(self, file_info: Dict):
        """保存文件元信息到JSON文件"""
        session_id = file_info.get('session_id')
        if session_id:
            user_folder = self.get_user_upload_folder(session_id)
            metadata_path = os.path.join(user_folder, f"{file_info['file_id']}_metadata.json")
        else:
            metadata_path = os.path.join(self.upload_folder, f"{file_info['file_id']}_metadata.json")
            
        try:
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(file_info, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存文件元信息失败: {e}")
    
    def preview_file(self, file_id: str, sheet_name: str = None, rows: int = 50, header_row: int = 0, session_id: str = None) -> Dict:
        """
        预览文件内容
        
        Args:
            file_id: 文件ID
            sheet_name: sheet名称（对Excel文件）
            rows: 预览行数
            header_row: 表头行号（0-based）
            session_id: 用户会话ID
            
        Returns:
            包含预览数据的字典
        """
        try:
            file_info = self.get_file_info(file_id, session_id)
            if not file_info:
                raise ValueError(f"文件不存在: {file_id}")
            
            file_path = file_info['file_path']
            extension = file_info['extension']
            
            # 读取数据
            if extension == 'csv':
                df = self._read_csv_with_encoding_detection(file_path, nrows=rows, header=header_row)
            elif extension in ['xlsx', 'xls']:
                if not sheet_name:
                    sheet_name = file_info['sheets'][0] if file_info['sheets'] else 0
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=rows, header=header_row)
            else:
                raise ValueError(f"不支持的文件类型: {extension}")
            
            # 准备预览数据
            preview_data = {
                'file_id': file_id,
                'filename': file_info['original_filename'],
                'sheet_name': sheet_name,
                'total_rows': len(df),
                'total_columns': len(df.columns),
                'columns': df.columns.tolist(),
                'data': df.fillna('').astype(str).to_dict('records'),
                'dtypes': df.dtypes.astype(str).to_dict()
            }
            
            return preview_data
            
        except Exception as e:
            logging.error(f"预览文件时出错: {e}")
            raise
    
    def _read_csv_with_encoding_detection(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        使用编码检测读取CSV文件
        """
        # 检测编码
        with open(file_path, 'rb') as f:
            raw_data = f.read(10000)  # 读取前10KB用于编码检测
            result = chardet.detect(raw_data)
            encoding = result['encoding']
        
        # 尝试读取CSV
        try:
            return pd.read_csv(file_path, encoding=encoding, **kwargs)
        except UnicodeDecodeError:
            # 如果检测的编码失败，尝试常用编码
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig', 'latin1']
            for enc in encodings:
                try:
                    return pd.read_csv(file_path, encoding=enc, **kwargs)
                except UnicodeDecodeError:
                    continue
            raise ValueError(f"无法确定CSV文件的编码格式: {file_path}")
    
    def read_full_file(self, file_id: str, sheet_name: str = None, header_row: int = 0, session_id: str = None) -> pd.DataFrame:
        """
        读取完整文件数据
        """
        file_info = self.get_file_info(file_id, session_id)
        if not file_info:
            raise ValueError(f"文件不存在: {file_id}")
        
        file_path = file_info['file_path']
        extension = file_info['extension']
        
        if extension == 'csv':
            return self._read_csv_with_encoding_detection(file_path, header=header_row)
        elif extension in ['xlsx', 'xls']:
            if not sheet_name:
                sheet_name = file_info['sheets'][0] if file_info['sheets'] else 0
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        else:
            raise ValueError(f"不支持的文件类型: {extension}")
    
    def save_result_file(self, df: pd.DataFrame, filename: str, format: str = 'xlsx', session_id: str = None) -> str:
        """
        保存结果文件
        
        Args:
            df: 要保存的DataFrame
            filename: 文件名（不含扩展名）
            format: 文件格式 ('xlsx' 或 'csv')
            session_id: 用户会话ID，用于文件隔离
            
        Returns:
            结果文件的路径
        """
        try:
            # 获取结果文件夹路径
            if session_id:
                results_folder = self.get_user_results_folder(session_id)
            else:
                results_folder = self.results_folder
                
            # 确保结果文件夹存在
            if not os.path.exists(results_folder):
                os.makedirs(results_folder, exist_ok=True)
                logging.info(f"创建results文件夹: {results_folder}")
            
            # 生成唯一的结果文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            result_filename = f"{filename}_{timestamp}.{format}"
            result_path = os.path.join(results_folder, result_filename)
            
            logging.info(f"准备保存结果文件: {result_path}")
            logging.info(f"DataFrame形状: {df.shape}")
            
            # 保存文件
            if format == 'xlsx':
                with pd.ExcelWriter(result_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='合并数据')
            elif format == 'csv':
                df.to_csv(result_path, index=False, encoding='utf-8-sig')
            else:
                raise ValueError(f"不支持的导出格式: {format}")
            
            # 验证文件是否成功创建
            if os.path.exists(result_path):
                file_size = os.path.getsize(result_path)
                logging.info(f"结果文件保存成功: {result_path}, 大小: {file_size} bytes")
            else:
                raise Exception(f"文件保存失败，文件不存在: {result_path}")
            
            return result_path
            
        except Exception as e:
            logging.error(f"保存结果文件时出错: {e}")
            logging.error(f"目标路径: {result_path if 'result_path' in locals() else 'unknown'}")
            logging.error(f"results_folder: {self.results_folder}")
            raise
    
    def cleanup_old_files(self, retention_days: int = 1):
        """
        清理过期的临时文件
        
        Args:
            retention_days: 文件保留天数
        """
        try:
            cutoff_time = datetime.now() - timedelta(days=retention_days)
            
            # 清理上传文件
            for filename in os.listdir(self.upload_folder):
                file_path = os.path.join(self.upload_folder, filename)
                if os.path.isfile(file_path):
                    file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if file_time < cutoff_time:
                        os.remove(file_path)
                        logging.info(f"已删除过期文件: {filename}")
            
            # 清理结果文件
            for filename in os.listdir(self.results_folder):
                file_path = os.path.join(self.results_folder, filename)
                if os.path.isfile(file_path):
                    file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if file_time < cutoff_time:
                        os.remove(file_path)
                        logging.info(f"已删除过期结果文件: {filename}")
                        
        except Exception as e:
            logging.error(f"清理文件时出错: {e}")
    
    def delete_file(self, file_id: str, session_id: str = None):
        """删除指定文件及其元信息"""
        try:
            file_info = self.get_file_info(file_id, session_id)
            if file_info:
                # 删除文件
                if os.path.exists(file_info['file_path']):
                    os.remove(file_info['file_path'])
                
                # 删除元信息文件
                if session_id:
                    user_folder = self.get_user_upload_folder(session_id)
                    metadata_path = os.path.join(user_folder, f"{file_id}_metadata.json")
                else:
                    metadata_path = os.path.join(self.upload_folder, f"{file_id}_metadata.json")
                
                if os.path.exists(metadata_path):
                    os.remove(metadata_path)
                
                logging.info(f"已删除文件: {file_id}")
                
        except Exception as e:
            logging.error(f"删除文件时出错: {e}")
    
    def get_file_list(self, session_id: str = None) -> List[Dict]:
        """获取文件列表，如果提供session_id则只返回该用户的文件"""
        files = []
        try:
            if session_id:
                # 获取用户专用文件夹中的文件
                user_folder = self.get_user_upload_folder(session_id)
                for filename in os.listdir(user_folder):
                    if filename.endswith('_metadata.json'):
                        file_id = filename.replace('_metadata.json', '')
                        file_info = self.get_file_info(file_id, session_id)
                        if file_info:
                            files.append(file_info)
            else:
                # 获取所有文件（向后兼容）
                for filename in os.listdir(self.upload_folder):
                    if filename.endswith('_metadata.json'):
                        file_id = filename.replace('_metadata.json', '')
                        file_info = self.get_file_info(file_id)
                        if file_info:
                            files.append(file_info)
        except Exception as e:
            logging.error(f"获取文件列表时出错: {e}")
        
        return sorted(files, key=lambda x: x['upload_time'], reverse=True)
    
    def clear_all_files(self, session_id: str = None) -> Dict[str, int]:
        """
        清理上传的文件和元数据
        
        Args:
            session_id: 用户会话ID，如果提供则只清理该用户的文件
        
        Returns:
            包含删除统计的字典
        """
        stats = {'files_deleted': 0, 'metadata_deleted': 0, 'errors': 0}
        
        try:
            if session_id:
                # 清理用户专用文件夹
                user_folder = self.get_user_upload_folder(session_id)
                if not os.path.exists(user_folder):
                    return stats
                folder_to_clean = user_folder
            else:
                # 清理所有文件（向后兼容）
                if not os.path.exists(self.upload_folder):
                    return stats
                folder_to_clean = self.upload_folder
                
            for filename in os.listdir(folder_to_clean):
                file_path = os.path.join(folder_to_clean, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        if filename.endswith('_metadata.json'):
                            stats['metadata_deleted'] += 1
                        else:
                            stats['files_deleted'] += 1
                except Exception as e:
                    logging.error(f"删除文件失败 {filename}: {e}")
                    stats['errors'] += 1
            
            logging.info(f"清理完成 - 文件: {stats['files_deleted']}, 元数据: {stats['metadata_deleted']}, 错误: {stats['errors']}")
            
        except Exception as e:
            logging.error(f"清理所有文件时出错: {e}")
            stats['errors'] += 1
            
        return stats
    
    def cleanup_old_files(self, retention_days: int = 1) -> Dict[str, int]:
        """
        清理过期文件
        
        Args:
            retention_days: 文件保留天数
            
        Returns:
            清理统计信息
        """
        stats = {'files_deleted': 0, 'metadata_deleted': 0, 'errors': 0}
        
        try:
            if not os.path.exists(self.upload_folder):
                return stats
                
            cutoff_time = time.time() - (retention_days * 24 * 3600)
            
            for filename in os.listdir(self.upload_folder):
                file_path = os.path.join(self.upload_folder, filename)
                
                try:
                    if os.path.isfile(file_path):
                        file_mtime = os.path.getmtime(file_path)
                        
                        if file_mtime < cutoff_time:
                            os.remove(file_path)
                            if filename.endswith('_metadata.json'):
                                stats['metadata_deleted'] += 1
                                logging.info(f"删除过期元数据: {filename}")
                            else:
                                stats['files_deleted'] += 1
                                logging.info(f"删除过期文件: {filename}")
                                
                except Exception as e:
                    logging.error(f"删除过期文件失败 {filename}: {e}")
                    stats['errors'] += 1
            
            if stats['files_deleted'] > 0 or stats['metadata_deleted'] > 0:
                logging.info(f"自动清理完成 - 文件: {stats['files_deleted']}, 元数据: {stats['metadata_deleted']}, 错误: {stats['errors']}")
            
        except Exception as e:
            logging.error(f"自动清理过期文件时出错: {e}")
            stats['errors'] += 1
            
        return stats