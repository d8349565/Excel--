import os
from datetime import timedelta

# 获取应用根目录的绝对路径
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    """应用配置类"""
    
    # 基本配置
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your-secret-key-here-change-in-production'
    
    # 用户认证配置
    ACCESS_PASSWORD = os.environ.get('ACCESS_PASSWORD') or '123456'
    ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD') or 'admin2025'  # 管理员密码
    SESSION_PERMANENT = False
    PERMANENT_SESSION_LIFETIME = timedelta(hours=24)  # 会话过期时间
    
    # 文件上传配置 - 使用绝对路径
    MAX_CONTENT_LENGTH = 500 * 1024 * 1024  # 500MB 总体上传限制
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    RESULTS_FOLDER = os.path.join(BASE_DIR, 'results')
    LOGS_FOLDER = os.path.join(BASE_DIR, 'logs')
    
    # 单文件大小限制 - 支持环境变量动态配置
    MAX_FILE_SIZE = int(os.environ.get('MAX_FILE_SIZE_MB', '100')) * 1024 * 1024  # 默认100MB
    
    # 支持的文件类型
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
    
    # 预览配置 - 支持环境变量动态配置
    DEFAULT_PREVIEW_ROWS = int(os.environ.get('DEFAULT_PREVIEW_ROWS', '50'))
    MAX_PREVIEW_ROWS = 200
    
    # 数据处理配置
    MAX_ERROR_SAMPLES = 100  # 最多显示错误样例数
    DEFAULT_HEADER_ROW = 1   # 默认表头行
    
    # 任务配置 - 支持环境变量动态配置
    MAX_CONCURRENT_TASKS = int(os.environ.get('MAX_CONCURRENT_TASKS', '5'))  # 最大并发任务数（支持多用户）
    TASK_TIMEOUT = 3600      # 任务超时时间（秒）
    
    # 文件清理配置 - 支持环境变量动态配置
    FILE_RETENTION_DAYS = int(os.environ.get('FILE_RETENTION_DAYS', '1'))  # 文件保留天数
    CLEANUP_INTERVAL_HOURS = 2  # 清理间隔（小时）
    
    # 日期格式
    DATE_OUTPUT_FORMAT = '%Y-%m-%d'
    COMMON_DATE_FORMATS = [
        '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
        '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
        '%d-%m-%Y', '%d.%m.%Y', '%Y%m%d'
    ]
    
    # 数值清洗配置
    CURRENCY_SYMBOLS = ['$', '€', '¥', '￥', '£', '₹', '₽']
    THOUSAND_SEPARATORS = [',', ' ', '.']
    
    # 导出配置
    DEFAULT_EXPORT_NAME_TEMPLATE = "merged_data_{timestamp}"
    
    @staticmethod
    def init_app(app):
        """初始化应用配置"""
        # 创建必要的目录
        folders = [
            app.config['UPLOAD_FOLDER'],
            app.config['RESULTS_FOLDER'],
            app.config['LOGS_FOLDER']
        ]
        
        for folder in folders:
            if not os.path.exists(folder):
                os.makedirs(folder)
    
    @staticmethod
    def get_user_folder(base_folder, session_id):
        """获取用户专用文件夹路径"""
        if session_id:
            user_folder = os.path.join(base_folder, session_id)
            if not os.path.exists(user_folder):
                os.makedirs(user_folder, exist_ok=True)
            return user_folder
        return base_folder

class DevelopmentConfig(Config):
    """开发环境配置"""
    DEBUG = True
    
class ProductionConfig(Config):
    """生产环境配置"""
    DEBUG = False
    
    # 生产环境应该使用更强的密钥
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'prod-secret-key-change-me'
    
    # 生产环境可以调整的参数
    MAX_CONCURRENT_TASKS = 2
    FILE_RETENTION_DAYS = 3

# 配置字典
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}