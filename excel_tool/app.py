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
            
            # 获取分页参数
            page = int(request.args.get('page', 1))
            page_size = min(int(request.args.get('page_size', 1000)), 5000)  # 限制最大1000行，防止性能问题
            show_all = request.args.get('show_all', 'false').lower() == 'true'
            
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
            import pandas as pd
            
            if show_all:
                # 显示全部数据（用于透视表分析）
                if filename.lower().endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                elif filename.lower().endswith('.csv'):
                    df = pd.read_csv(file_path, encoding='utf-8-sig')
                else:
                    return jsonify({
                        'success': False,
                        'message': '不支持的文件格式'
                    }), 400
            else:
                # 分页读取（用于预览显示）
                if filename.lower().endswith('.xlsx'):
                    # Excel文件先读取少量数据获取总行数
                    df_sample = pd.read_excel(file_path, nrows=0)  # 只读取列名
                    df_full = pd.read_excel(file_path)  # 读取全部数据用于计算总行数
                    total_rows = len(df_full)
                    
                    # 计算分页
                    start_row = (page - 1) * page_size
                    end_row = start_row + page_size
                    df = df_full.iloc[start_row:end_row]
                    
                elif filename.lower().endswith('.csv'):
                    # CSV文件分页处理
                    df_full = pd.read_csv(file_path, encoding='utf-8-sig')
                    total_rows = len(df_full)
                    
                    start_row = (page - 1) * page_size
                    end_row = start_row + page_size
                    df = df_full.iloc[start_row:end_row]
                else:
                    return jsonify({
                        'success': False,
                        'message': '不支持的文件格式'
                    }), 400
            
            # 转换为JSON格式
            data = {
                'columns': df.columns.tolist(),
                'data': df.fillna('').astype(str).values.tolist(),
                'current_rows': len(df),
                'total_rows': len(df) if show_all else total_rows,
                'page': page if not show_all else 1,
                'page_size': page_size if not show_all else len(df),
                'total_pages': max(1, (total_rows + page_size - 1) // page_size) if not show_all else 1,
                'show_all': show_all,
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
                
                # 映射聚合函数
                agg_func_mapping = {
                    'sum': 'sum',
                    'count': 'count',
                    'avg': 'mean', 
                    'mean': 'mean',
                    'min': 'min',
                    'max': 'max'
                }
                agg_func = agg_func_mapping.get(aggregation, 'sum')
                
                # 创建透视表
                pivot_table = pd.pivot_table(
                    df, 
                    values=values, 
                    index=index, 
                    columns=columns, 
                    aggfunc=agg_func,
                    fill_value=0
                )
                
                # 生成统计信息
                stats = {
                    'total_records': len(df),
                    'pivot_rows': len(pivot_table.index) if hasattr(pivot_table, 'index') else 1,
                    'pivot_columns': len(pivot_table.columns) if hasattr(pivot_table, 'columns') else 1,
                    'aggregation_method': aggregation,
                    'fields_used': {
                        'row_fields': row_fields,
                        'column_fields': column_fields,
                        'value_fields': value_fields
                    }
                }
                
                # 计算数值统计
                if hasattr(pivot_table, 'values'):
                    values_flat = pivot_table.values.flatten()
                    values_flat = values_flat[~pd.isna(values_flat)]
                    if len(values_flat) > 0:
                        stats['data_summary'] = {
                            'min_value': float(values_flat.min()),
                            'max_value': float(values_flat.max()),
                            'sum_value': float(values_flat.sum()),
                            'avg_value': float(values_flat.mean()),
                            'non_zero_count': int((values_flat != 0).sum())
                        }
                
                # 转换透视表为可序列化格式
                if hasattr(pivot_table, 'index') and hasattr(pivot_table, 'columns'):
                    # 多维透视表
                    pivot_data = {
                        'index': [str(idx) for idx in pivot_table.index.tolist()],
                        'columns': [str(col) for col in pivot_table.columns.tolist()],
                        'data': pivot_table.fillna(0).values.tolist(),
                        'index_name': pivot_table.index.name or 'Index',
                        'columns_name': pivot_table.columns.name or 'Columns'
                    }
                else:
                    # 简单聚合结果
                    if isinstance(pivot_table, pd.Series):
                        pivot_data = {
                            'index': [str(idx) for idx in pivot_table.index.tolist()],
                            'columns': [values],
                            'data': [[val] for val in pivot_table.fillna(0).values.tolist()],
                            'index_name': pivot_table.index.name or row_fields[0] if row_fields else 'Index',
                            'columns_name': values
                        }
                    else:
                        # 单一值结果
                        pivot_data = {
                            'index': ['总计'],
                            'columns': [values],
                            'data': [[float(pivot_table) if pd.notna(pivot_table) else 0]],
                            'index_name': 'Summary',
                            'columns_name': values
                        }
                
                return jsonify({
                    'success': True,
                    'pivot_data': pivot_data,
                    'stats': stats
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

    @app.route('/api/results/pivot/export', methods=['POST'])
    @login_required
    def export_pivot_table():
        """导出透视表API"""
        try:
            data = request.get_json()
            filename = data.get('filename')
            row_fields = data.get('row_fields', [])
            column_fields = data.get('column_fields', [])
            value_fields = data.get('value_fields', [])
            aggregation = data.get('aggregation', 'sum')
            export_format = data.get('format', 'xlsx')  # xlsx 或 csv
            
            if not filename or not value_fields:
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
            
            # 生成透视表
            try:
                # 数据清洗：确保数据质量
                df = df.copy()  # 创建副本避免修改原始数据
                
                # 检查并清理字段名中的特殊字符
                df.columns = [str(col).strip() for col in df.columns]
                
                # 清理数据：处理无限值和非数值
                df = df.replace([float('inf'), float('-inf')], 0)
                
                index = row_fields if row_fields else None
                columns = column_fields if column_fields else None
                values = value_fields[0]
                
                # 确保值字段存在且为数值类型
                if values not in df.columns:
                    return jsonify({
                        'success': False,
                        'message': f'值字段 "{values}" 不存在'
                    }), 400
                
                # 尝试将值字段转换为数值类型
                try:
                    df[values] = pd.to_numeric(df[values], errors='coerce')
                    df[values] = df[values].fillna(0)  # 将无法转换的值设为0
                except Exception as e:
                    app.logger.warning(f"值字段转换警告: {e}")
                
                # 映射聚合函数
                agg_func_mapping = {
                    'sum': 'sum',
                    'count': 'count',
                    'avg': 'mean',
                    'mean': 'mean',
                    'min': 'min',
                    'max': 'max'
                }
                agg_func = agg_func_mapping.get(aggregation, 'sum')
                
                # 创建透视表
                pivot_table = pd.pivot_table(
                    df,
                    values=values,
                    index=index,
                    columns=columns,
                    aggfunc=agg_func,
                    fill_value=0,
                    dropna=False  # 不删除包含NaN的行
                )
                
                # 生成导出文件名 - 使用英文避免兼容性问题
                base_name = os.path.splitext(safe_filename)[0]
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                export_filename = f"{base_name}_pivot_{timestamp}.{export_format}"
                export_path = os.path.join(user_results_folder, export_filename)
                
                # 创建包含透视表和统计信息的完整报告
                if export_format.lower() == 'xlsx':
                    try:
                        # 使用 ExcelWriter 创建多工作表的Excel文件，移除不兼容的选项
                        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                            # 写入透视表
                            if isinstance(pivot_table, pd.DataFrame):
                                # 重置索引以包含行字段作为列
                                pivot_df = pivot_table.reset_index()
                                # 处理NaN值和数据类型
                                pivot_df = pivot_df.fillna(0)  # 将NaN替换为0
                                # 确保所有列名都是字符串
                                pivot_df.columns = [str(col) for col in pivot_df.columns]
                                # 确保数值列为数值类型
                                for col in pivot_df.columns:
                                    if pivot_df[col].dtype == 'object':
                                        try:
                                            pivot_df[col] = pd.to_numeric(pivot_df[col], errors='ignore')
                                        except:
                                            pass
                                pivot_df.to_excel(writer, sheet_name='PivotTable', index=False)
                            elif isinstance(pivot_table, pd.Series):
                                # 转换Series为DataFrame
                                pivot_df = pivot_table.reset_index()
                                pivot_df.columns = [str(pivot_df.columns[0]), str(values)]
                                pivot_df = pivot_df.fillna(0)  # 处理NaN值
                                pivot_df.to_excel(writer, sheet_name='PivotTable', index=False)
                            else:
                                # 处理单一值的情况
                                pivot_df = pd.DataFrame({
                                    'Field': [str(values)],
                                    'Value': [float(pivot_table) if pd.notna(pivot_table) else 0]
                                })
                                pivot_df.to_excel(writer, sheet_name='PivotTable', index=False)
                            
                            # 创建统计信息工作表
                            stats_data = {
                                'Item': [
                                    'Source File', 'Total Records', 'Pivot Rows', 'Pivot Columns',
                                    'Aggregation', 'Row Fields', 'Column Fields', 'Value Fields', 'Generated Time'
                                ],
                                'Value': [
                                    str(filename),
                                    str(len(df)),
                                    str(len(pivot_table.index) if hasattr(pivot_table, 'index') else 1),
                                    str(len(pivot_table.columns) if hasattr(pivot_table, 'columns') else 1),
                                    str(aggregation),
                                    ', '.join(row_fields) if row_fields else 'None',
                                    ', '.join(column_fields) if column_fields else 'None',
                                    ', '.join(value_fields),
                                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                ]
                            }
                            
                            stats_df = pd.DataFrame(stats_data)
                            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
                            
                            # 写入完整的原始数据
                            full_data_df = df.copy()
                            # 处理原始数据中的问题值
                            full_data_df = full_data_df.fillna('')  # 将NaN替换为空字符串
                            # 确保列名为字符串
                            full_data_df.columns = [str(col) for col in full_data_df.columns]
                            # 转换所有列为字符串以避免类型问题
                            for col in full_data_df.columns:
                                try:
                                    full_data_df[col] = full_data_df[col].astype(str)
                                except:
                                    full_data_df[col] = full_data_df[col].fillna('').astype(str)
                            full_data_df.to_excel(writer, sheet_name='SourceData', index=False)
                    
                    except Exception as excel_error:
                        app.logger.error(f"Excel文件创建失败: {excel_error}")
                        # 如果Excel创建失败，尝试创建CSV作为备选
                        csv_filename = export_filename.replace('.xlsx', '.csv')
                        csv_path = os.path.join(user_results_folder, csv_filename)
                        
                        if isinstance(pivot_table, pd.DataFrame):
                            pivot_df = pivot_table.reset_index().fillna(0)
                        elif isinstance(pivot_table, pd.Series):
                            pivot_df = pivot_table.reset_index()
                            pivot_df.columns = [str(pivot_df.columns[0]), str(values)]
                            pivot_df = pivot_df.fillna(0)
                        else:
                            pivot_df = pd.DataFrame({'Field': [str(values)], 'Value': [float(pivot_table) if pd.notna(pivot_table) else 0]})
                        
                        pivot_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                        export_filename = csv_filename
                        export_path = csv_path
                        
                elif export_format.lower() == 'csv':
                    # CSV格式只能包含透视表数据
                    if isinstance(pivot_table, pd.DataFrame):
                        pivot_df = pivot_table.reset_index()
                    elif isinstance(pivot_table, pd.Series):
                        pivot_df = pivot_table.reset_index()
                        pivot_df.columns = [pivot_df.columns[0], values]
                    else:
                        pivot_df = pd.DataFrame({'结果': [pivot_table]})
                    
                    pivot_df.to_csv(export_path, index=False, encoding='utf-8-sig')
                
                # 记录导出日志
                user_logger.log_file_download(
                    filename=export_filename,
                    session_id=session_id
                )
                
                return jsonify({
                    'success': True,
                    'message': '透视表导出成功',
                    'filename': export_filename,
                    'download_url': url_for('download_file', filename=export_filename)
                })
                
            except Exception as export_error:
                logging.error(f"导出透视表失败: {str(export_error)}")
                return jsonify({
                    'success': False,
                    'message': f'导出透视表失败: {str(export_error)}'
                }), 500
            
        except Exception as e:
            logging.error(f"导出透视表失败: {str(e)}")
            return jsonify({
                'success': False,
                'message': f'导出透视表失败: {str(e)}'
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
            sort_field = data.get('sort_field', 'none')  # 排序字段: none, x_axis, y_axis
            sort_direction = data.get('sort_direction', 'asc')  # 排序方向: asc, desc
            sort_type = data.get('sort_type', 'alphabetic')  # 排序方式: alphabetic, numeric, date, custom
            custom_sort_order = data.get('custom_sort_order', [])  # 自定义排序顺序
            
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
            import numpy as np
            from datetime import datetime
            
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
                
                # 应用排序
                if sort_field != 'none':
                    if sort_field == 'x_axis':
                        sort_column = x_field
                    elif sort_field == 'y_axis':
                        sort_column = y_field
                    else:
                        sort_column = x_field
                    
                    # 根据排序方式进行排序
                    if sort_type == 'custom' and custom_sort_order:
                        # 自定义排序
                        def custom_sort_key(x):
                            try:
                                return custom_sort_order.index(str(x))
                            except ValueError:
                                return len(custom_sort_order)  # 未找到的项放在最后
                        
                        grouped['sort_key'] = grouped[sort_column].apply(custom_sort_key)
                        grouped = grouped.sort_values('sort_key', ascending=(sort_direction == 'asc'))
                        grouped = grouped.drop('sort_key', axis=1)
                        
                    elif sort_type == 'numeric':
                        # 数值排序
                        try:
                            grouped[sort_column] = pd.to_numeric(grouped[sort_column], errors='coerce')
                            grouped = grouped.sort_values(sort_column, ascending=(sort_direction == 'asc'))
                        except:
                            # 如果转换失败，回退到字母排序
                            grouped = grouped.sort_values(sort_column, ascending=(sort_direction == 'asc'))
                            
                    elif sort_type == 'date':
                        # 日期排序
                        try:
                            grouped[sort_column] = pd.to_datetime(grouped[sort_column], errors='coerce')
                            grouped = grouped.sort_values(sort_column, ascending=(sort_direction == 'asc'))
                            grouped[sort_column] = grouped[sort_column].astype(str)
                        except:
                            # 如果转换失败，回退到字母排序
                            grouped = grouped.sort_values(sort_column, ascending=(sort_direction == 'asc'))
                            
                    else:
                        # 字母排序（默认）
                        grouped = grouped.sort_values(sort_column, ascending=(sort_direction == 'asc'))
                
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