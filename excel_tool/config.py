import os
from datetime import timedelta

# 获取应用根目录的绝对路径
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    """应用配置类"""
    
    # 基本配置
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your-secret-key-here-change-in-production'
    
    # 文件上传配置 - 使用绝对路径
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB 总体上传限制
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    RESULTS_FOLDER = os.path.join(BASE_DIR, 'results')
    LOGS_FOLDER = os.path.join(BASE_DIR, 'logs')
    
    # 单文件大小限制
    MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
    
    # 支持的文件类型
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
    
    # 预览配置
    DEFAULT_PREVIEW_ROWS = 50
    MAX_PREVIEW_ROWS = 200
    
    # 数据处理配置
    MAX_ERROR_SAMPLES = 100  # 最多显示错误样例数
    DEFAULT_HEADER_ROW = 1   # 默认表头行
    
    # 任务配置
    MAX_CONCURRENT_TASKS = 1  # 最大并发任务数
    TASK_TIMEOUT = 3600      # 任务超时时间（秒）
    
    # 文件清理配置
    FILE_RETENTION_DAYS = 1  # 文件保留天数
    CLEANUP_INTERVAL_HOURS = 6  # 清理间隔（小时）
    
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