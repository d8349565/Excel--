from flask import Flask, request, render_template, redirect, url_for, jsonify, send_file, flash, session
import os
import logging
from datetime import datetime
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

def login_required(f):
    """装饰器：检查用户是否已登录"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'authenticated' not in session or not session['authenticated']:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def get_session_id():
    """获取或创建会话ID"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return session['session_id']

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
                flash('登录成功，欢迎使用Excel/CSV汇总工具！', 'success')
                return redirect(url_for('index'))
            else:
                flash('密码错误，请重新输入', 'error')
        
        return render_template('login.html')
    
    @app.route('/logout')
    def logout():
        """用户登出"""
        session.clear()
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
                    file_info = file_manager.save_uploaded_file(file, file.filename, session_id)
                    file_ids.append(file_info['file_id'])
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
            
            return jsonify({
                'success': True,
                'data': preview_data
            })
            
        except Exception as e:
            return jsonify({
                'success': False,
                'error': str(e)
            }), 400
    
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
                file_info = file_manager.get_file_info(file_id, session_id)
                if not file_info:
                    continue
                
                # 读取数据
                df = file_manager.read_full_file(file_id, sheet_name, header_row, session_id)
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
                if cleaning_options.get('clean_numeric', False):
                    user_column_types = cleaning_options.get('column_types', {})
                    configured_df = processor.clean_numeric_data(configured_df, user_column_types=user_column_types)
                
                if cleaning_options.get('parse_dates', False):
                    user_column_types = cleaning_options.get('column_types', {})
                    configured_df = processor.parse_dates(configured_df, user_column_types=user_column_types)
                
                if cleaning_options.get('remove_duplicates', False):
                    keep_strategy = cleaning_options.get('keep_strategy', 'first')
                    configured_df = processor.remove_duplicates(configured_df, keep=keep_strategy)
                
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
            file_manager.delete_file(file_id, session_id)
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
            stats = file_manager.clear_all_files(session_id)
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
        return jsonify({
            'success': True,
            'status': {
                'queue_size': task_manager.get_queue_size(),
                'running_tasks': task_manager.get_running_tasks(),
                'total_tasks': len(task_manager.tasks),
                'uptime': str(datetime.now())
            }
        })
    
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
        
        # 设置文件日志
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
        ))
        file_handler.setLevel(logging.INFO)
        app.logger.addHandler(file_handler)
        
        app.logger.setLevel(logging.INFO)
        app.logger.info('Excel汇总工具启动')

def startup_cleanup(app, file_manager):
    """应用启动时清理upload文件夹中的所有文件"""
    try:
        app.logger.info("开始清理启动时的临时文件...")
        
        # 清理所有用户的上传文件（不分session）
        upload_folder = app.config['UPLOAD_FOLDER']
        if not os.path.exists(upload_folder):
            app.logger.info("上传文件夹不存在，无需清理")
            return
        
        total_deleted = 0
        error_count = 0
        
        # 遍历upload文件夹中的所有内容
        for item in os.listdir(upload_folder):
            item_path = os.path.join(upload_folder, item)
            
            try:
                if os.path.isfile(item_path):
                    # 删除文件
                    os.remove(item_path)
                    total_deleted += 1
                elif os.path.isdir(item_path):
                    # 删除用户文件夹及其内容
                    import shutil
                    shutil.rmtree(item_path)
                    total_deleted += 1
                    app.logger.info(f"删除用户文件夹: {item}")
            except Exception as e:
                error_count += 1
                app.logger.error(f"删除启动文件失败 {item}: {e}")
        
        if total_deleted > 0:
            app.logger.info(f"启动清理完成 - 删除了 {total_deleted} 个文件/文件夹")
        else:
            app.logger.info("启动时没有需要清理的文件")
            
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
    app.run(debug=True, host='0.0.0.0', port=5000)