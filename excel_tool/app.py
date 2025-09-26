from flask import Flask, request, render_template, redirect, url_for, jsonify, send_file, flash, session
import os
import logging
from datetime import datetime, timezone, timedelta
from werkzeug.utils import secure_filename
import atexit
import threading
import time
import uuid
from functools import wraps

from config import config
from file_manager import FileManager
from data_processor import DataProcessor
from task_manager import task_manager, create_merge_task_handler
from user_logger import UserLogger

def login_required(f):
    """装饰器：检查用户是否已登录"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'authenticated' not in session or not session['authenticated']:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    """装饰器：检查用户是否具有管理员权限"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'authenticated' not in session or not session['authenticated']:
            return redirect(url_for('login'))
        if 'is_admin' not in session or not session['is_admin']:
            flash('需要管理员权限才能访问此页面', 'error')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

def get_session_id():
    """获取或创建会话ID"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return session['session_id']

def format_file_size(size_bytes):
    """格式化文件大小"""
    if size_bytes == 0:
        return "0B"
    size_names = ["B", "KB", "MB", "GB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"

def create_app(config_name=None):
    """创建Flask应用实例"""
    app = Flask(__name__)
    
    # 配置应用
    config_name = config_name or os.environ.get('FLASK_CONFIG') or 'default'
    app.config.from_object(config[config_name])
    config[config_name].init_app(app)
    
    # 设置日志
    setup_logging(app)
    
    # 记录重要路径信息
    app.logger.info(f"应用根目录: {os.path.abspath(os.path.dirname(__file__))}")
    app.logger.info(f"上传文件夹: {app.config['UPLOAD_FOLDER']}")
    app.logger.info(f"结果文件夹: {app.config['RESULTS_FOLDER']}")
    app.logger.info(f"日志文件夹: {app.config['LOGS_FOLDER']}")
    
    # 初始化组件
    file_manager = FileManager(app)
    data_processor = DataProcessor(app.config)
    user_logger = UserLogger(app.config)
    
    # 将组件添加到app实例
    app.file_manager = file_manager
    app.data_processor = data_processor
    app.user_logger = user_logger
    
    # 启动任务管理器
    task_manager.max_workers = app.config.get('MAX_CONCURRENT_TASKS', 1)
    task_manager.task_timeout = app.config.get('TASK_TIMEOUT', 3600)
    task_manager.start()
    
    # 注册任务处理器
    merge_handler = create_merge_task_handler(file_manager, data_processor)
    task_manager.register_handler('merge_data', merge_handler)
    
    # 启动清理线程
    start_cleanup_thread(app, file_manager)
    
    # 应用启动时清理upload文件夹
    startup_cleanup(app, file_manager)
    
    # 注册关闭处理
    atexit.register(lambda: task_manager.stop())
    
    # 添加模板函数
    @app.template_global()
    def get_operation_name(operation):
        """获取操作类型的中文名称"""
        operation_names = {
            'user_login': '用户登录',
            'user_logout': '用户登出',
            'file_upload': '文件上传',
            'file_delete': '文件删除',
            'file_preview': '文件预览',
            'merge_task_submit': '提交合并任务',
            'merge_task_complete': '合并完成',
            'file_download': '文件下载',
            'clear_all_files': '清空文件'
        }
        return operation_names.get(operation, operation)
    
    # 路由定义
    @app.route('/login', methods=['GET', 'POST'])
    def login():
        """用户登录"""
        if request.method == 'POST':
            password = request.form.get('password', '').strip()
            
            if password == app.config['ACCESS_PASSWORD']:
                session['authenticated'] = True
                session['session_id'] = str(uuid.uuid4())
                session.permanent = True
                
                # 记录登录日志
                user_logger.log_login(session['session_id'])
                
                flash('登录成功，欢迎使用Excel/CSV汇总工具！', 'success')
                return redirect(url_for('index'))
            else:
                flash('密码错误，请重新输入', 'error')
        
        return render_template('login.html')
    
    @app.route('/admin/login', methods=['GET', 'POST'])
    def admin_login():
        """管理员登录"""
        if request.method == 'POST':
            password = request.form.get('password', '').strip()
            
            if password == app.config['ADMIN_PASSWORD']:
                session['is_admin'] = True
                session.permanent = True
                flash('管理员登录成功！', 'success')
                
                # 如果有重定向目标，跳转到目标页面
                next_page = request.args.get('next')
                if next_page:
                    return redirect(next_page)
                return redirect(url_for('user_logs'))
            else:
                flash('管理员密码错误，请重新输入', 'error')
        
        return render_template('admin_login.html')
    
    @app.route('/logout')
    def logout():
        """用户登出"""
        session_id = session.get('session_id')
        is_admin = session.get('is_admin', False)
        
        # 记录登出日志
        if session_id:
            user_logger.log_logout(session_id)
        
        session.clear()
        
        if is_admin:
            flash('已退出管理员模式', 'info')
        else:
            flash('您已安全退出', 'info')
            
        return redirect(url_for('login'))
    
    @app.route('/')
    @login_required
    def index():
        """首页"""
        session_id = get_session_id()
        files = file_manager.get_file_list(session_id)
        return render_template('index.html', files=files)
    
    @app.route('/upload', methods=['POST'])
    @login_required
    def upload_files():
        """处理文件上传"""
        try:
            session_id = get_session_id()
            uploaded_files = request.files.getlist('files')
            
            if not uploaded_files or all(f.filename == '' for f in uploaded_files):
                flash('请选择要上传的文件', 'error')
                return redirect(url_for('index'))
            
            file_ids = []
            
            for file in uploaded_files:
                if file.filename == '':
                    continue
                
                try:
                    # 获取文件大小
                    file.seek(0, 2)  # 移动到文件末尾
                    file_size = file.tell()
                    file.seek(0)  # 重置到开头
                    
                    file_info = file_manager.save_uploaded_file(file, file.filename, session_id)
                    file_ids.append(file_info['file_id'])
                    
                    # 记录文件上传日志
                    user_logger.log_file_upload(
                        filename=file.filename,
                        file_size=file_size,
                        file_id=file_info['file_id'],
                        session_id=session_id
                    )
                    
                    flash(f'文件 {file.filename} 上传成功', 'success')
                except Exception as e:
                    flash(f'文件 {file.filename} 上传失败: {str(e)}', 'error')
            
            if file_ids:
                return redirect(url_for('configure', file_ids=','.join(file_ids)))
            else:
                return redirect(url_for('index'))
                
        except Exception as e:
            flash(f'上传失败: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    @app.route('/configure')
    @login_required
    def configure():
        """配置页面"""
        session_id = get_session_id()
        file_ids_str = request.args.get('file_ids', '')
        
        if not file_ids_str:
            flash('没有指定要配置的文件', 'error')
            return redirect(url_for('index'))
        
        file_ids = file_ids_str.split(',')
        files_info = []
        
        for file_id in file_ids:
            file_info = file_manager.get_file_info(file_id, session_id)
            if file_info:
                files_info.append(file_info)
        
        if not files_info:
            flash('指定的文件不存在', 'error')
            return redirect(url_for('index'))
        
        return render_template('configure.html', files=files_info)
    
    @app.route('/api/preview/<file_id>')
    @login_required
    def preview_file(file_id):
        """文件预览API"""
        try:
            session_id = get_session_id()
            sheet_name = request.args.get('sheet_name')
            rows = int(request.args.get('rows', app.config['DEFAULT_PREVIEW_ROWS']))
            header_row = int(request.args.get('header_row', 0))
            
            # 限制预览行数
            rows = min(rows, app.config['MAX_PREVIEW_ROWS'])
            
            # 验证文件是否属于当前用户
            file_info = file_manager.get_file_info(file_id, session_id)
            if not file_info:
                return jsonify({
                    'success': False,
                    'error': '文件不存在或无权访问'
                }), 404
            
            preview_data = file_manager.preview_file(file_id, sheet_name, rows, header_row, session_id)
            
            # 记录文件预览日志
            user_logger.log_file_preview(
                filename=file_info.get('original_filename', 'unknown'),
                file_id=file_id,
                sheet_name=sheet_name,
                session_id=session_id
            )
            
            return jsonify({
                'success': True,
                'data': preview_data
            })
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 400
    
    @app.route('/api/file_sheets/<file_id>')
    @login_required
    def get_file_sheets(file_id):
        """获取文件的工作表列表"""
        try:
            session_id = get_session_id()
            
            # 验证文件权限
            file_info = app.file_manager.get_file_info(file_id, session_id)
            if not file_info:
                return jsonify({'error': '文件不存在或无权访问'}), 404
            
            # 获取工作表列表
            sheets = app.file_manager.get_sheet_names(file_id, session_id)
            
            return jsonify({
                'success': True,
                'sheets': sheets
            })
            
        except Exception as e:
            app.logger.error(f"获取工作表列表失败: {str(e)}")
            return jsonify({'error': f'获取工作表列表失败: {str(e)}'}), 500

    @app.route('/api/preview_cell_value', methods=['POST'])
    @login_required
    def preview_cell_value():
        """预览指定单元格的值"""
        try:
            session_id = get_session_id()
            data = request.get_json()
            
            file_id = data.get('file_id')
            sheet_name = data.get('sheet_name')
            cell_address = data.get('cell_address')
            
            if not all([file_id, sheet_name, cell_address]):
                return jsonify({'error': '缺少必要参数'}), 400
            
            # 验证文件权限
            file_info = app.file_manager.get_file_info(file_id, session_id)
            if not file_info:
                return jsonify({'error': '文件不存在或无权访问'}), 404
            
            # 获取单元格值
            cell_value = app.file_manager.read_cell_value_by_address(file_id, sheet_name, cell_address, session_id)
            
            return jsonify({
                'success': True,
                'value': str(cell_value) if cell_value is not None else ''
            })
            
        except Exception as e:
            app.logger.error(f"预览单元格值失败: {str(e)}")
            return jsonify({'error': f'预览单元格值失败: {str(e)}'}), 500

    @app.route('/api/detect_columns/<file_id>')
    @login_required
    def detect_columns(file_id):
        """检测列类型API"""
        try:
            session_id = get_session_id()
            sheet_name = request.args.get('sheet_name')
            header_row = int(request.args.get('header_row', 0))
            
            # 验证文件权限
            file_info = file_manager.get_file_info(file_id, session_id)
            if not file_info:
                return jsonify({
                    'success': False,
                    'error': '文件不存在或无权访问'
                }), 404
            
            # 读取部分数据用于类型检测
            df = file_manager.read_full_file(file_id, sheet_name, header_row, session_id)
            
            # 标准化列名
            df = data_processor.standardize_column_names(df)
            
            # 检测列类型
            column_types = data_processor.detect_column_types(df)
            
            return jsonify({
                'success': True,
                'column_types': column_types,
                'columns': df.columns.tolist()
            })
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 400
    
    @app.route('/get_merged_columns', methods=['POST'])
    @login_required
    def get_merged_columns():
        """获取合并后的列信息"""
        try:
            session_id = get_session_id()
            data = request.get_json()
            file_configs = data.get('file_configs', [])
            
            if not file_configs:
                return jsonify({'success': False, 'error': '没有文件配置'})
            
            # 准备数据框列表
            dataframes = []
            for config in file_configs:
                file_id = config['file_id']
                sheet_name = config.get('sheet_name', 0)
                header_row = config.get('header_row', 0)
                
                # 验证文件权限
                file_info = file_manager.get_file_info(file_id, session_id)
                if not file_info:
                    continue
                
                # 读取数据
                df = file_manager.read_full_file(file_id, sheet_name, header_row, session_id)
                if df is not None:
                    dataframes.append((df, file_id))
            
            if not dataframes:
                return jsonify({'success': False, 'error': '无法读取文件'})
            
            # 合并数据框获取列信息
            processor = DataProcessor()
            try:
                # 使用外连接合并以获取所有列
                merged_df = processor.merge_dataframes(dataframes, 'outer')
                
                # 分析每列的数据类型
                columns_info = []
                for col in merged_df.columns:
                    # 获取列的样本数据（非空值）
                    sample_data = merged_df[col].dropna().head(100)
                    
                    # 尝试推断数据类型
                    suggested_type = 'auto'
                    if len(sample_data) > 0:
                        # 检查是否包含数值模式
                        if processor._contains_numeric_pattern(sample_data):
                            suggested_type = 'number'
                        # 检查是否包含日期模式
                        elif processor._contains_date_pattern(sample_data):
                            suggested_type = 'date'
                        else:
                            suggested_type = 'text'
                    
                    columns_info.append({
                        'name': col,
                        'suggested_type': suggested_type,
                        'sample_values': sample_data.head(5).tolist()
                    })
                
                return jsonify({
                    'success': True,
                    'columns': columns_info
                })
                
            except Exception as e:
                return jsonify({'success': False, 'error': f'分析列信息失败: {str(e)}'})
            
        except Exception as e:
            logging.error(f"获取列信息失败: {e}")
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/preview_merge', methods=['POST'])
    @login_required
    def preview_merge():
        """预览合并结果"""
        try:
            session_id = get_session_id()
            data = request.get_json()
            file_configs = data.get('file_configs', [])
            cleaning_options = data.get('cleaning_options', {})
            
            if not file_configs:
                return jsonify({'success': False, 'error': '没有文件配置'})
            
            # 准备数据框列表
            dataframes = []
            for config in file_configs:
                file_id = config['file_id']
                sheet_name = config.get('sheet_name', 0)
                header_row = config.get('header_row', 0)
                source_name = config.get('source_name', f'文件{len(dataframes)+1}')
                
                # 验证文件权限
                file_info = app.file_manager.get_file_info(file_id, session_id)
                if not file_info:
                    continue
                
                # 读取数据
                df = app.file_manager.read_full_file(file_id, sheet_name, header_row, session_id)
                if df is not None:
                    # 限制预览数据量，每个文件最多取10行
                    preview_df = df.head(10)
                    dataframes.append((preview_df, source_name))
            
            if not dataframes:
                return jsonify({'success': False, 'error': '无法读取文件'})
            
            # 处理数据（使用预览数据但应用完整的处理流程）
            processor = DataProcessor()
            try:
                # 合并数据
                merged_df = processor.merge_dataframes(
                    dataframes, 
                    cleaning_options.get('merge_strategy', 'outer')
                )
                
                # 应用列配置（顺序、显示、重命名）
                configured_df = processor.apply_column_configuration(merged_df, cleaning_options)
                
                # 应用数据清洗选项（如果启用）
                # 1. 去除空行
                if cleaning_options.get('remove_empty_rows', False):
                    key_columns = cleaning_options.get('key_columns')
                    if key_columns:
                        # 处理两种情况：字符串或列表
                        if isinstance(key_columns, str):
                            key_columns = [col.strip() for col in key_columns.split(',') if col.strip()]
                        elif isinstance(key_columns, list):
                            key_columns = [str(col).strip() for col in key_columns if str(col).strip()]
                        else:
                            key_columns = None
                    configured_df = processor.remove_empty_rows(configured_df, key_columns)
                
                # 2. 数值清洗
                if cleaning_options.get('clean_numeric', False):
                    numeric_columns = cleaning_options.get('numeric_columns')
                    user_column_types = cleaning_options.get('column_types', {})
                    configured_df = processor.clean_numeric_data(configured_df, numeric_columns, user_column_types=user_column_types)
                
                # 3. 日期解析
                if cleaning_options.get('parse_dates', False):
                    date_columns = cleaning_options.get('date_columns')
                    user_column_types = cleaning_options.get('column_types', {})
                    configured_df = processor.parse_dates(configured_df, date_columns, user_column_types=user_column_types)
                
                # 4. 去重
                if cleaning_options.get('remove_duplicates', False):
                    duplicate_columns = cleaning_options.get('duplicate_columns')
                    keep_strategy = cleaning_options.get('keep_strategy', 'first')
                    configured_df = processor.remove_duplicates(configured_df, duplicate_columns, keep_strategy)
                
                # 5. 处理固定单元格数据提取（预览模式）
                if cleaning_options.get('fixed_cells_rules'):
                    configured_df = processor.extract_fixed_cells_data(
                        configured_df, 
                        cleaning_options.get('fixed_cells_rules'), 
                        app.file_manager, 
                        session_id
                    )
                
                # 应用数据清洗选项完成
                
                # 限制预览结果，最多20行
                preview_result = configured_df.head(20)
                
                # 处理NaN值，转换为可序列化的格式
                preview_result = preview_result.fillna('')  # 将NaN替换为空字符串
                
                # 转换为可序列化的格式
                result_data = {
                    'success': True,
                    'preview_data': {
                        'columns': preview_result.columns.tolist(),
                        'data': preview_result.values.tolist(),
                        'total_rows': len(configured_df),
                        'preview_rows': len(preview_result),
                        'strategy': cleaning_options.get('merge_strategy', 'outer')
                    },
                    'stats': {
                        'input_files': len(dataframes),
                        'total_columns': len(configured_df.columns),
                        'estimated_total_rows': len(configured_df),
                        'applied_configurations': {
                            'column_order': bool(cleaning_options.get('column_order')),
                            'column_names': bool(cleaning_options.get('column_names')),
                            'hidden_columns': bool(cleaning_options.get('hidden_columns')),
                            'clean_numeric': cleaning_options.get('clean_numeric', False),
                            'parse_dates': cleaning_options.get('parse_dates', False),
                            'remove_duplicates': cleaning_options.get('remove_duplicates', False)
                        }
                    }
                }
                
                return jsonify(result_data)
                
            except Exception as e:
                return jsonify({'success': False, 'error': f'预览合并失败: {str(e)}'})
            
        except Exception as e:
            logging.error(f"预览合并失败: {e}")
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/submit_task', methods=['POST'])
    @login_required
    def submit_task():
        """提交合并任务"""
        try:
            session_id = get_session_id()
            data = request.get_json()
            
            file_configs = data.get('file_configs', [])
            cleaning_options = data.get('cleaning_options', {})
            export_options = data.get('export_options', {})
            
            if not file_configs:
                return jsonify({
                    'success': False,
                    'error': '没有指定要处理的文件'
                }), 400
            
            # 验证所有文件都属于当前用户
            for config in file_configs:
                file_id = config.get('file_id')
                if file_id:
                    file_info = file_manager.get_file_info(file_id, session_id)
                    if not file_info:
                        return jsonify({
                            'success': False,
                            'error': f'文件 {file_id} 不存在或无权访问'
                        }), 403
            
            # 提交任务
            task_id = task_manager.submit_task('merge_data', {
                'file_configs': file_configs,
                'cleaning_options': cleaning_options,
                'export_options': export_options,
                'session_id': session_id
            })
            
            # 记录任务提交日志
            user_logger.log_merge_task_submit(
                file_count=len(file_configs),
                task_id=task_id,
                export_format=export_options.get('format', 'xlsx'),
                session_id=session_id
            )
            
            return jsonify({
                'success': True,
                'task_id': task_id
            })
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 400
    
    @app.route('/task/<task_id>')
    @login_required
    def task_status(task_id):
        """任务状态页面"""
        task = task_manager.get_task(task_id)
        
        if not task:
            flash('指定的任务不存在', 'error')
            return redirect(url_for('index'))
        
        return render_template('task.html', task=task.to_dict())
    
    @app.route('/api/task/<task_id>')
    def api_task_status(task_id):
        """任务状态API"""
        task_data = task_manager.get_task_status(task_id)
        
        if not task_data:
            return jsonify({
                'success': False,
                'error': '任务不存在'
            }), 404
        
        return jsonify({
            'success': True,
            'task': task_data
        })
    
    @app.route('/download/<filename>')
    @login_required
    def download_file(filename):
        """下载结果文件"""
        try:
            session_id = get_session_id()
            # 安全检查文件名
            safe_filename = secure_filename(filename)
            
            # 使用用户专用的结果文件夹
            user_results_folder = file_manager.get_user_results_folder(session_id)
            file_path = os.path.join(user_results_folder, safe_filename)
            
            # 如果用户文件夹中没有，再尝试主文件夹（向后兼容）
            if not os.path.exists(file_path):
                file_path = os.path.join(app.config['RESULTS_FOLDER'], safe_filename)
            
            if not os.path.exists(file_path):
                flash('请求的文件不存在或已被删除', 'error')
                return redirect(url_for('index'))
            
            # 记录文件下载日志
            user_logger.log_file_download(
                filename=filename,
                session_id=session_id
            )
            
            return send_file(file_path, as_attachment=True, download_name=filename)
            
        except Exception as e:
            logging.error(f"下载文件失败 - 文件: {filename}, 错误: {str(e)}")
            flash('文件下载失败，请稍后重试', 'error')
            return redirect(url_for('index'))
    
    @app.route('/api/delete_file/<file_id>', methods=['DELETE'])
    @login_required
    def delete_file(file_id):
        """删除文件API"""
        try:
            session_id = get_session_id()
            
            # 获取文件信息用于日志记录
            file_info = file_manager.get_file_info(file_id, session_id)
            filename = file_info.get('original_filename', 'unknown') if file_info else 'unknown'
            
            # 删除文件
            file_manager.delete_file(file_id, session_id)
            
            # 记录删除日志
            user_logger.log_file_delete(
                filename=filename,
                file_id=file_id,
                session_id=session_id
            )
            
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 400
    
    @app.route('/api/clear_all_files', methods=['POST'])
    @login_required
    def clear_all_files():
        """清理所有文件API"""
        try:
            session_id = get_session_id()
            logging.info(f"开始清理用户 {session_id} 的所有文件")
            
            # 获取文件数量用于日志记录
            files = file_manager.get_file_list(session_id)
            files_count = len(files)
            logging.info(f"用户 {session_id} 有 {files_count} 个文件需要清理")
            
            # 如果没有文件，直接返回
            if files_count == 0:
                return jsonify({
                    'success': True, 
                    'message': '没有文件需要清理',
                    'stats': {'files_deleted': 0, 'metadata_deleted': 0, 'errors': 0}
                })
            
            # 执行清理
            stats = file_manager.clear_all_files(session_id)
            logging.info(f"清理完成: {stats}")
            
            # 记录清空文件日志
            user_logger.log_clear_all_files(
                files_count=files_count,
                session_id=session_id
            )
            
            return jsonify({
                'success': True, 
                'message': f'已清理 {stats["files_deleted"]} 个文件和 {stats["metadata_deleted"]} 个元数据记录',
                'stats': stats
            })
        except Exception as e:
            logging.error(f"清理所有文件失败: {e}")
            return jsonify({'success': False, 'error': '清理失败，请稍后重试'}), 500
    
    @app.route('/api/system_status')
    @login_required
    def system_status():
        """系统状态API"""
        # 使用北京时区显示时间
        beijing_tz = timezone(timedelta(hours=8))
        current_time = datetime.now(beijing_tz).strftime('%Y-%m-%d %H:%M:%S %Z')
        
        return jsonify({
            'success': True,
            'status': {
                'queue_size': task_manager.get_queue_size(),
                'running_tasks': task_manager.get_running_tasks(),
                'total_tasks': len(task_manager.tasks),
                'uptime': current_time
            }
        })
    
    @app.route('/logs')
    @login_required
    @admin_required
    def user_logs():
        """用户日志查看页面（仅管理员）"""
        session_id = get_session_id()
        
        # 获取查询参数
        operation_filter = request.args.get('operation')
        limit = min(int(request.args.get('limit', 50)), 200)  # 最多200条
        
        # 管理员可以查看所有用户的日志，不传session_id
        logs = user_logger.get_user_logs(
            session_id=None,  # 管理员查看所有日志
            limit=limit,
            operation_filter=operation_filter
        )
        
        # 获取统计信息（所有用户）
        stats = user_logger.get_operation_stats(session_id=None, days=7)
        
        return render_template('user_logs.html', logs=logs, stats=stats)
    
    @app.route('/api/user_logs')
    @login_required
    @admin_required
    def api_user_logs():
        """用户日志API（仅管理员）"""
        session_id = get_session_id()
        
        # 获取查询参数
        operation_filter = request.args.get('operation')
        limit = min(int(request.args.get('limit', 50)), 200)
        
        # 管理员可以查看所有用户的日志
        logs = user_logger.get_user_logs(
            session_id=None,  # 管理员查看所有日志
            limit=limit,
            operation_filter=operation_filter
        )
        
        return jsonify({
            'success': True,
            'logs': logs,
            'total': len(logs)
        })

    @app.route('/results')
    @login_required
    def results_page():
        """结果管理页面"""
        return render_template('results.html')

    @app.route('/api/results')
    @login_required
    def api_results():
        """获取处理结果列表API"""
        try:
            session_id = get_session_id()
            user_results_folder = file_manager.get_user_results_folder(session_id)
            
            results = []
            if os.path.exists(user_results_folder):
                for filename in os.listdir(user_results_folder):
                    file_path = os.path.join(user_results_folder, filename)
                    if os.path.isfile(file_path) and filename.endswith(('.xlsx', '.csv')):
                        stat = os.stat(file_path)
                        results.append({
                            'id': filename.replace('.', '_'),  # 用于DOM ID
                            'filename': filename,
                            'size': stat.st_size,
                            'size_formatted': format_file_size(stat.st_size),
                            'created_time': datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                            'modified_time': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                        })
            
            # 按修改时间降序排列
            results.sort(key=lambda x: x['modified_time'], reverse=True)
            
            return jsonify({
                'success': True,
                'results': results
            })
            
        except Exception as e:
            logging.error(f"获取结果列表失败: {str(e)}")
            return jsonify({
                'success': False,
                'message': '获取结果列表失败'
            }), 500

    @app.route('/api/results/delete', methods=['POST'])
    @login_required
    def delete_results():
        """批量删除结果文件API"""
        try:
            data = request.get_json()
            filenames = data.get('filenames', [])
            
            if not filenames:
                return jsonify({
                    'success': False,
                    'message': '未选择要删除的文件'
                }), 400
            
            session_id = get_session_id()
            user_results_folder = file_manager.get_user_results_folder(session_id)
            
            deleted_files = []
            failed_files = []
            
            for filename in filenames:
                try:
                    safe_filename = secure_filename(filename)
                    file_path = os.path.join(user_results_folder, safe_filename)
                    
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        deleted_files.append(filename)
                        
                        # 记录删除日志
                        user_logger.log_file_delete(
                            filename=filename,
                            session_id=session_id
                        )
                    else:
                        failed_files.append(f"{filename} (文件不存在)")
                        
                except Exception as e:
                    failed_files.append(f"{filename} ({str(e)})")
            
            return jsonify({
                'success': True,
                'deleted': deleted_files,
                'failed': failed_files,
                'message': f'成功删除 {len(deleted_files)} 个文件'
            })
            
        except Exception as e:
            logging.error(f"批量删除文件失败: {str(e)}")
            return jsonify({
                'success': False,
                'message': '删除文件失败'
            }), 500

    @app.route('/api/results/preview/<filename>')
    @login_required
    def preview_result_data(filename):
        """预览结果文件数据API"""
        try:
            session_id = get_session_id()
            safe_filename = secure_filename(filename)
            
            user_results_folder = file_manager.get_user_results_folder(session_id)
            file_path = os.path.join(user_results_folder, safe_filename)
            
            if not os.path.exists(file_path):
                return jsonify({
                    'success': False,
                    'message': '文件不存在'
                }), 404
            
            # 使用数据处理器读取文件
            processor = DataProcessor()
            
            # 根据文件类型读取数据
            if filename.lower().endswith('.xlsx'):
                import pandas as pd
                df = pd.read_excel(file_path, nrows=100)  # 只读取前100行用于预览
            elif filename.lower().endswith('.csv'):
                import pandas as pd
                df = pd.read_csv(file_path, nrows=100, encoding='utf-8-sig')
            else:
                return jsonify({
                    'success': False,
                    'message': '不支持的文件格式'
                }), 400
            
            # 转换为JSON格式
            data = {
                'columns': df.columns.tolist(),
                'data': df.fillna('').astype(str).values.tolist(),
                'total_rows': len(df),
                'dtypes': {col: str(dtype) for col, dtype in df.dtypes.items()}
            }
            
            return jsonify({
                'success': True,
                'data': data
            })
            
        except Exception as e:
            logging.error(f"预览结果文件失败: {filename}, 错误: {str(e)}")
            return jsonify({
                'success': False,
                'message': f'预览文件失败: {str(e)}'
            }), 500

    @app.route('/api/results/pivot', methods=['POST'])
    @login_required
    def generate_pivot_table():
        """生成数据透视表API"""
        try:
            data = request.get_json()
            filename = data.get('filename')
            row_fields = data.get('row_fields', [])
            column_fields = data.get('column_fields', [])
            value_fields = data.get('value_fields', [])
            aggregation = data.get('aggregation', 'sum')
            
            if not filename:
                return jsonify({
                    'success': False,
                    'message': '缺少文件名参数'
                }), 400
            
            session_id = get_session_id()
            safe_filename = secure_filename(filename)
            
            user_results_folder = file_manager.get_user_results_folder(session_id)
            file_path = os.path.join(user_results_folder, safe_filename)
            
            if not os.path.exists(file_path):
                return jsonify({
                    'success': False,
                    'message': '文件不存在'
                }), 404
            
            # 读取完整数据用于透视分析
            import pandas as pd
            if filename.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif filename.lower().endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                return jsonify({
                    'success': False,
                    'message': '不支持的文件格式'
                }), 400
            
            # 生成透视表
            try:
                # 构建透视表
                index = row_fields if row_fields else None
                columns = column_fields if column_fields else None
                values = value_fields[0] if value_fields else None
                
                if not values:
                    return jsonify({
                        'success': False,
                        'message': '必须指定值字段'
                    }), 400
                
                # 创建透视表
                pivot_table = pd.pivot_table(
                    df, 
                    values=values, 
                    index=index, 
                    columns=columns, 
                    aggfunc=aggregation,
                    fill_value=0
                )
                
                # 转换为JSON格式
                pivot_data = {
                    'index': pivot_table.index.tolist() if hasattr(pivot_table.index, 'tolist') else [str(pivot_table.index)],
                    'columns': pivot_table.columns.tolist() if hasattr(pivot_table.columns, 'tolist') else [str(pivot_table.columns)],
                    'data': pivot_table.values.tolist()
                }
                
                return jsonify({
                    'success': True,
                    'pivot_data': pivot_data
                })
                
            except Exception as pivot_error:
                return jsonify({
                    'success': False,
                    'message': f'生成透视表失败: {str(pivot_error)}'
                }), 500
            
        except Exception as e:
            logging.error(f"生成透视表失败: {str(e)}")
            return jsonify({
                'success': False,
                'message': f'生成透视表失败: {str(e)}'
            }), 500

    @app.route('/api/results/chart-data', methods=['POST'])
    @login_required
    def generate_chart_data():
        """生成图表数据API"""
        try:
            data = request.get_json()
            filename = data.get('filename')
            x_field = data.get('x_field')
            y_field = data.get('y_field')
            chart_type = data.get('chart_type', 'bar')
            aggregation = data.get('aggregation', 'sum')
            
            if not all([filename, x_field, y_field]):
                return jsonify({
                    'success': False,
                    'message': '缺少必要参数'
                }), 400
            
            session_id = get_session_id()
            safe_filename = secure_filename(filename)
            
            user_results_folder = file_manager.get_user_results_folder(session_id)
            file_path = os.path.join(user_results_folder, safe_filename)
            
            if not os.path.exists(file_path):
                return jsonify({
                    'success': False,
                    'message': '文件不存在'
                }), 404
            
            # 读取数据
            import pandas as pd
            if filename.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif filename.lower().endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                return jsonify({
                    'success': False,
                    'message': '不支持的文件格式'
                }), 400
            
            # 检查字段是否存在
            if x_field not in df.columns or y_field not in df.columns:
                return jsonify({
                    'success': False,
                    'message': '指定的字段不存在'
                }), 400
            
            # 生成图表数据
            try:
                # 按x字段分组，对y字段进行聚合
                agg_func = aggregation if aggregation in ['sum', 'count', 'mean', 'min', 'max'] else 'sum'
                if agg_func == 'avg':
                    agg_func = 'mean'
                    
                grouped = df.groupby(x_field)[y_field].agg(agg_func).reset_index()
                
                chart_data = {
                    'labels': grouped[x_field].astype(str).tolist(),
                    'datasets': [{
                        'label': f'{y_field} ({aggregation})',
                        'data': grouped[y_field].tolist(),
                        'backgroundColor': 'rgba(54, 162, 235, 0.2)',
                        'borderColor': 'rgba(54, 162, 235, 1)',
                        'borderWidth': 1
                    }]
                }
                
                return jsonify({
                    'success': True,
                    'chart_data': chart_data
                })
                
            except Exception as chart_error:
                return jsonify({
                    'success': False,
                    'message': f'生成图表数据失败: {str(chart_error)}'
                }), 500
            
        except Exception as e:
            logging.error(f"生成图表数据失败: {str(e)}")
            return jsonify({
                'success': False,
                'message': f'生成图表数据失败: {str(e)}'
            }), 500
    
    # 错误处理
    @app.errorhandler(404)
    def not_found(error):
        return render_template('error.html', 
                             error_code=404, 
                             error_message='页面未找到'), 404
    
    @app.errorhandler(500)
    def internal_error(error):
        return render_template('error.html', 
                             error_code=500, 
                             error_message='服务器内部错误'), 500
    
    @app.errorhandler(413)
    def request_entity_too_large(error):
        return render_template('error.html', 
                             error_code=413, 
                             error_message='上传文件过大'), 413
    
    return app

def setup_logging(app):
    """设置日志"""
    if not app.debug:
        if not os.path.exists(app.config['LOGS_FOLDER']):
            os.makedirs(app.config['LOGS_FOLDER'])
        
        log_file = os.path.join(app.config['LOGS_FOLDER'], 'app.log')
        
        # 设置北京时区
        beijing_tz = timezone(timedelta(hours=8))
        
        # 设置文件日志
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        file_handler.setLevel(logging.INFO)
        app.logger.addHandler(file_handler)
        
        app.logger.setLevel(logging.INFO)
        app.logger.info('Excel汇总工具启动')

def startup_cleanup(app, file_manager):
    """应用启动时清理过期的临时文件"""
    try:
        app.logger.info("开始清理启动时的过期临时文件...")
        
        # 清理过期文件（1天前），而不是所有文件
        retention_days = app.config.get('FILE_RETENTION_DAYS', 1)
        
        upload_folder = app.config['UPLOAD_FOLDER']
        if not os.path.exists(upload_folder):
            app.logger.info("上传文件夹不存在，无需清理")
            return
        
        total_deleted = 0
        error_count = 0
        cutoff_time = time.time() - (retention_days * 24 * 3600)
        
        # 遍历upload文件夹中的所有内容
        for item in os.listdir(upload_folder):
            item_path = os.path.join(upload_folder, item)
            
            try:
                # 检查文件/文件夹的修改时间
                item_mtime = os.path.getmtime(item_path)
                
                if item_mtime < cutoff_time:
                    if os.path.isfile(item_path):
                        # 删除过期文件
                        os.remove(item_path)
                        total_deleted += 1
                        app.logger.info(f"删除过期文件: {item}")
                    elif os.path.isdir(item_path):
                        # 删除过期用户文件夹及其内容
                        import shutil
                        shutil.rmtree(item_path)
                        total_deleted += 1
                        app.logger.info(f"删除过期用户文件夹: {item}")
                else:
                    app.logger.debug(f"保留未过期的文件/文件夹: {item}")
                    
            except Exception as e:
                error_count += 1
                app.logger.error(f"检查启动文件失败 {item}: {e}")
        
        if total_deleted > 0:
            app.logger.info(f"启动清理完成 - 删除了 {total_deleted} 个过期文件/文件夹")
        else:
            app.logger.info("启动时没有过期文件需要清理")
            
        if error_count > 0:
            app.logger.warning(f"启动清理过程中遇到 {error_count} 个错误")
            
    except Exception as e:
        app.logger.error(f"启动清理失败: {e}")

def start_cleanup_thread(app, file_manager):
    """启动文件清理线程"""
    def cleanup_worker():
        while True:
            try:
                with app.app_context():
                    # 清理过期上传文件
                    retention_days = app.config.get('FILE_RETENTION_DAYS', 1)
                    file_manager.cleanup_old_files(retention_days)
                    
                    # 清理过期结果文件
                    cleanup_old_results(app, retention_days)
                    
                    # 清理过期任务
                    task_manager.cleanup_old_tasks(24)
                
                # 等待下次清理
                cleanup_interval = app.config.get('CLEANUP_INTERVAL_HOURS', 6) * 3600
                time.sleep(cleanup_interval)
                
            except Exception as e:
                app.logger.error(f'清理线程出错: {e}')
                time.sleep(3600)  # 出错后等待1小时再重试
    
    cleanup_thread = threading.Thread(target=cleanup_worker, name='CleanupWorker')
    cleanup_thread.daemon = True
    cleanup_thread.start()
    
    app.logger.info('文件清理线程已启动')

def cleanup_old_results(app, retention_days: int = 1):
    """清理过期的结果文件"""
    try:
        results_folder = app.config['RESULTS_FOLDER']
        if not os.path.exists(results_folder):
            return
            
        cutoff_time = time.time() - (retention_days * 24 * 3600)
        deleted_count = 0
        
        for filename in os.listdir(results_folder):
            file_path = os.path.join(results_folder, filename)
            
            try:
                if os.path.isfile(file_path):
                    file_mtime = os.path.getmtime(file_path)
                    
                    if file_mtime < cutoff_time:
                        os.remove(file_path)
                        deleted_count += 1
                        app.logger.info(f"删除过期结果文件: {filename}")
                        
            except Exception as e:
                app.logger.error(f"删除过期结果文件失败 {filename}: {e}")
        
        if deleted_count > 0:
            app.logger.info(f"清理了 {deleted_count} 个过期结果文件")
            
    except Exception as e:
        app.logger.error(f"清理结果文件时出错: {e}")

if __name__ == '__main__':
    app = create_app()
    app.run(debug=False, host='0.0.0.0', port=5000)