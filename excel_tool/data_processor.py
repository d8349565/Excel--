import pandas as pd
import numpy as np
import re
import logging
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any
import warnings

warnings.filterwarnings('ignore')

class DataProcessor:
    """数据处理器，负责数据清洗、合并等核心功能"""
    
    def __init__(self, config=None):
        self.config = config or {}
        self.currency_symbols = self.config.get('CURRENCY_SYMBOLS', ['$', '€', '¥', '￥', '£', '₹', '₽'])
        self.thousand_separators = self.config.get('THOUSAND_SEPARATORS', [',', ' ', '.'])
        self.date_formats = self.config.get('COMMON_DATE_FORMATS', [
            '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
            '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
            '%d-%m-%Y', '%d.%m.%Y', '%Y%m%d'
        ])
        self.date_output_format = self.config.get('DATE_OUTPUT_FORMAT', '%Y-%m-%d')
        self.max_error_samples = self.config.get('MAX_ERROR_SAMPLES', 100)
        
        # 错误收集
        self.errors = []
        self.processing_stats = {
            'input_rows': 0,
            'output_rows': 0,
            'duplicates_removed': 0,
            'errors_count': 0,
            'columns_processed': 0
        }
    
    def reset_stats(self):
        """重置处理统计"""
        self.errors = []
        self.processing_stats = {
            'input_rows': 0,
            'output_rows': 0,
            'duplicates_removed': 0,
            'errors_count': 0,
            'columns_processed': 0
        }
    
    def standardize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        标准化列名
        
        Args:
            df: 输入DataFrame
            
        Returns:
            列名标准化后的DataFrame
        """
        try:
            new_columns = []
            for col in df.columns:
                # 转换为字符串并去除首尾空白
                col_str = str(col).strip()
                
                # 替换特殊字符为下划线
                col_str = re.sub(r'[^\w\u4e00-\u9fff]+', '_', col_str)
                
                # 去除连续的下划线
                col_str = re.sub(r'_+', '_', col_str)
                
                # 去除首尾下划线
                col_str = col_str.strip('_')
                
                # 如果列名为空，使用默认名称
                if not col_str:
                    col_str = f'column_{len(new_columns)}'
                
                new_columns.append(col_str)
            
            df.columns = new_columns
            return df
            
        except Exception as e:
            logging.error(f"标准化列名时出错: {e}")
            raise
    
    def remove_empty_rows(self, df: pd.DataFrame, key_columns: List[str] = None) -> pd.DataFrame:
        """
        去除空行
        
        Args:
            df: 输入DataFrame
            key_columns: 关键列，如果指定，则只要这些列都为空就删除行
            
        Returns:
            去除空行后的DataFrame
        """
        try:
            original_rows = len(df)
            
            if key_columns:
                # 只检查关键列
                valid_key_cols = [col for col in key_columns if col in df.columns]
                if valid_key_cols:
                    mask = df[valid_key_cols].notna().any(axis=1)
                    df = df[mask]
                else:
                    logging.warning(f"指定的关键列不存在: {key_columns}")
            else:
                # 检查所有列
                df = df.dropna(how='all')
            
            removed_rows = original_rows - len(df)
            if removed_rows > 0:
                logging.info(f"去除了 {removed_rows} 个空行")
            
            return df
            
        except Exception as e:
            logging.error(f"去除空行时出错: {e}")
            raise
    
    def clean_numeric_data(self, df: pd.DataFrame, numeric_columns: List[str] = None, 
                          user_column_types: Dict[str, str] = None) -> pd.DataFrame:
        """
        清洗数值数据（去除千分位符、货币符号等）
        
        Args:
            df: 输入DataFrame
            numeric_columns: 要处理的数值列，如果为None则自动检测
            user_column_types: 用户指定的列类型字典
            
        Returns:
            数值清洗后的DataFrame
        """
        try:
            df_result = df.copy()
            
            if numeric_columns is None:
                # 自动检测可能的数值列
                numeric_columns = []
                for col in df.columns:
                    # 优先使用用户指定的列类型
                    if user_column_types and col in user_column_types:
                        if user_column_types[col] == 'numeric':
                            numeric_columns.append(col)
                    else:
                        # 检查列中是否包含数值模式
                        if self._contains_numeric_pattern(df[col]):
                            numeric_columns.append(col)
            
            for col in numeric_columns:
                if col not in df.columns:
                    continue
                
                # 跳过用户明确指定为非数值类型的列
                if user_column_types and col in user_column_types:
                    if user_column_types[col] != 'numeric':
                        continue
                
                original_values = df_result[col].copy()
                cleaned_values = []
                
                for idx, value in enumerate(original_values):
                    try:
                        if pd.isna(value):
                            cleaned_values.append(value)
                            continue
                        
                        cleaned_value = self._clean_single_numeric_value(str(value))
                        cleaned_values.append(cleaned_value)
                        
                    except Exception as e:
                        # 记录错误样例
                        if len(self.errors) < self.max_error_samples:
                            self.errors.append({
                                'type': 'numeric_conversion',
                                'column': col,
                                'row': idx,
                                'original_value': str(value),
                                'error': str(e)
                            })
                        
                        cleaned_values.append(value)  # 保留原值
                        self.processing_stats['errors_count'] += 1
                
                df_result[col] = cleaned_values
            
            return df_result
            
        except Exception as e:
            logging.error(f"清洗数值数据时出错: {e}")
            raise
    
    def _contains_numeric_pattern(self, series: pd.Series) -> bool:
        """检查系列是否包含数值模式"""
        # 先检查pandas数据类型
        if series.dtype in [np.int64, np.float64, np.int32, np.float32]:
            return True
        
        # 获取非空值样本
        sample_values = series.dropna().astype(str).head(50)
        if len(sample_values) == 0:
            return False
        
        # 检查是否大部分值都能转换为数值
        numeric_count = 0
        for value in sample_values:
            value = str(value).strip()
            # 去除货币符号和千分位符号后检查
            cleaned_value = value
            for symbol in self.currency_symbols:
                cleaned_value = cleaned_value.replace(symbol, '')
            cleaned_value = cleaned_value.replace(',', '').replace(' ', '')
            
            # 尝试转换为数值
            try:
                float(cleaned_value)
                numeric_count += 1
            except ValueError:
                # 检查是否是百分比
                if cleaned_value.endswith('%'):
                    try:
                        float(cleaned_value[:-1])
                        numeric_count += 1
                    except ValueError:
                        pass
        
        # 如果超过80%的值可以转换为数值，则认为是数值列
        return numeric_count / len(sample_values) >= 0.8
    
    def _clean_single_numeric_value(self, value: str) -> float:
        """清洗单个数值"""
        if not value or pd.isna(value):
            return np.nan
        
        # 去除货币符号
        for symbol in self.currency_symbols:
            value = value.replace(symbol, '')
        
        # 去除空格
        value = value.strip()
        
        # 处理千分位分隔符
        # 首先检查是否是欧洲格式（小数点用逗号）
        if ',' in value and '.' in value:
            # 同时包含逗号和点，判断哪个是千分位分隔符
            last_comma = value.rfind(',')
            last_dot = value.rfind('.')
            
            if last_comma > last_dot:
                # 逗号在后面，是小数分隔符
                value = value.replace('.', '').replace(',', '.')
            else:
                # 点在后面，是小数分隔符
                value = value.replace(',', '')
        elif ',' in value and value.count(',') == 1 and len(value.split(',')[1]) <= 2:
            # 只有一个逗号且后面最多2位数字，可能是小数分隔符
            value = value.replace(',', '.')
        else:
            # 去除千分位逗号
            value = value.replace(',', '')
        
        # 尝试转换为浮点数
        try:
            return float(value)
        except ValueError:
            raise ValueError(f"无法转换为数值: {value}")
    
    def parse_dates(self, df: pd.DataFrame, date_columns: List[str] = None, 
                   user_column_types: Dict[str, str] = None) -> pd.DataFrame:
        """
        解析日期数据
        
        Args:
            df: 输入DataFrame
            date_columns: 要处理的日期列，如果为None则自动检测
            user_column_types: 用户指定的列类型字典
            
        Returns:
            日期解析后的DataFrame
        """
        try:
            df_result = df.copy()
            
            if date_columns is None:
                # 自动检测可能的日期列
                date_columns = []
                for col in df.columns:
                    # 优先使用用户指定的列类型
                    if user_column_types and col in user_column_types:
                        if user_column_types[col] == 'date':
                            date_columns.append(col)
                    else:
                        if self._contains_date_pattern(df[col]):
                            date_columns.append(col)
            
            for col in date_columns:
                if col not in df.columns:
                    continue
                
                # 跳过用户明确指定为非日期类型的列
                if user_column_types and col in user_column_types:
                    if user_column_types[col] != 'date':
                        continue
                
                original_values = df_result[col].copy()
                parsed_values = []
                
                for idx, value in enumerate(original_values):
                    try:
                        if pd.isna(value):
                            parsed_values.append(value)
                            continue
                        
                        parsed_date = self._parse_single_date(str(value))
                        if parsed_date:
                            parsed_values.append(parsed_date.strftime(self.date_output_format))
                        else:
                            parsed_values.append(value)  # 保留原值
                            
                    except Exception as e:
                        # 记录错误样例
                        if len(self.errors) < self.max_error_samples:
                            self.errors.append({
                                'type': 'date_parsing',
                                'column': col,
                                'row': idx,
                                'original_value': str(value),
                                'error': str(e)
                            })
                        
                        parsed_values.append(value)  # 保留原值
                        self.processing_stats['errors_count'] += 1
                
                df_result[col] = parsed_values
            
            return df_result
            
        except Exception as e:
            logging.error(f"解析日期时出错: {e}")
            raise
    
    def _contains_date_pattern(self, series: pd.Series) -> bool:
        """检查系列是否包含日期模式"""
        # 简单的日期模式检测
        date_patterns = [
            r'\\d{4}[-/]\\d{1,2}[-/]\\d{1,2}',  # YYYY-MM-DD 或 YYYY/MM/DD
            r'\\d{1,2}[-/]\\d{1,2}[-/]\\d{4}',  # DD-MM-YYYY 或 MM/DD/YYYY
            r'\\d{8}',  # YYYYMMDD
        ]
        
        sample_values = series.dropna().astype(str).head(20)
        for pattern in date_patterns:
            if sample_values.str.contains(pattern, regex=True, na=False).any():
                return True
        return False
    
    def _parse_single_date(self, date_str: str) -> Optional[datetime]:
        """解析单个日期字符串"""
        if not date_str or pd.isna(date_str):
            return None
        
        # 清理日期字符串
        date_str = str(date_str).strip()
        
        # 尝试pandas的智能解析
        try:
            return pd.to_datetime(date_str, infer_datetime_format=True)
        except:
            pass
        
        # 尝试预定义格式
        for fmt in self.date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    def remove_duplicates(self, df: pd.DataFrame, subset: List[str] = None, keep: str = 'first') -> pd.DataFrame:
        """
        去除重复行
        
        Args:
            df: 输入DataFrame
            subset: 用于判断重复的列，如果为None则使用所有列
            keep: 保留策略 ('first', 'last', False)
            
        Returns:
            去重后的DataFrame
        """
        try:
            original_rows = len(df)
            
            df_result = df.drop_duplicates(subset=subset, keep=keep)
            
            duplicates_removed = original_rows - len(df_result)
            self.processing_stats['duplicates_removed'] += duplicates_removed
            
            if duplicates_removed > 0:
                logging.info(f"去除了 {duplicates_removed} 个重复行")
            
            return df_result
            
        except Exception as e:
            logging.error(f"去除重复行时出错: {e}")
            raise
    
    def merge_dataframes(self, dataframes: List[Tuple[pd.DataFrame, str]], merge_strategy: str = 'outer') -> pd.DataFrame:
        """
        合并多个DataFrame
        
        Args:
            dataframes: DataFrame列表，每个元素是(df, source_name)的元组
            merge_strategy: 合并策略 ('outer', 'inner')
            
        Returns:
            合并后的DataFrame
        """
        try:
            if not dataframes:
                raise ValueError("没有可合并的数据")
            
            # 标准化所有DataFrame的列名
            standardized_dfs = []
            for df, source_name in dataframes:
                df_std = self.standardize_column_names(df.copy())
                # 添加数据源列
                df_std['数据源'] = source_name
                standardized_dfs.append(df_std)
                self.processing_stats['input_rows'] += len(df_std)
            
            # 使用pandas concat进行合并
            if merge_strategy == 'outer':
                result_df = pd.concat(standardized_dfs, ignore_index=True, sort=False)
            elif merge_strategy == 'inner':
                # 获取所有DataFrame的共同列
                common_columns = set(standardized_dfs[0].columns)
                for df in standardized_dfs[1:]:
                    common_columns &= set(df.columns)
                
                # 只保留共同列
                filtered_dfs = [df[list(common_columns)] for df in standardized_dfs]
                result_df = pd.concat(filtered_dfs, ignore_index=True, sort=False)
            else:
                raise ValueError(f"不支持的合并策略: {merge_strategy}")
            
            self.processing_stats['output_rows'] = len(result_df)
            self.processing_stats['columns_processed'] = len(result_df.columns)
            
            logging.info(f"成功合并 {len(dataframes)} 个数据源，共 {len(result_df)} 行，{len(result_df.columns)} 列")
            
            return result_df
            
        except Exception as e:
            logging.error(f"合并DataFrame时出错: {e}")
            raise
    
    def process_data(self, dataframes: List[Tuple[pd.DataFrame, str]], 
                    cleaning_options: Dict[str, Any], file_manager=None, session_id: str = None,
                    file_configs: List[Dict] = None) -> pd.DataFrame:
        """
        完整的数据处理流程
        
        Args:
            dataframes: DataFrame列表
            cleaning_options: 清洗选项
            file_manager: 文件管理器实例
            session_id: 会话ID
            file_configs: 文件配置列表，包含文件ID和数据源名称映射
            
        Returns:
            处理后的DataFrame
        """
        try:
            self.reset_stats()
            
            # 获取合并策略
            merge_strategy = cleaning_options.get('merge_strategy', 'outer')
            logging.info(f"使用合并策略: {merge_strategy}")
            
            # 合并数据
            merged_df = self.merge_dataframes(dataframes, merge_strategy)
            
            # 应用清洗步骤
            result_df = merged_df
            
            # 应用列配置（顺序、显示、重命名）
            result_df = self.apply_column_configuration(result_df, cleaning_options)
            
            # 1. 去除空行
            if cleaning_options.get('remove_empty_rows', False):
                key_columns = cleaning_options.get('key_columns')
                result_df = self.remove_empty_rows(result_df, key_columns)
            
            # 2. 数值清洗
            if cleaning_options.get('clean_numeric', False):
                numeric_columns = cleaning_options.get('numeric_columns')
                user_column_types = cleaning_options.get('column_types', {})
                result_df = self.clean_numeric_data(result_df, numeric_columns, user_column_types=user_column_types)
            
            # 3. 日期解析
            if cleaning_options.get('parse_dates', False):
                date_columns = cleaning_options.get('date_columns')
                user_column_types = cleaning_options.get('column_types', {})
                result_df = self.parse_dates(result_df, date_columns, user_column_types=user_column_types)
            
            # 4. 去重
            if cleaning_options.get('remove_duplicates', False):
                duplicate_columns = cleaning_options.get('duplicate_columns')
                keep_strategy = cleaning_options.get('keep_strategy', 'first')
                result_df = self.remove_duplicates(result_df, duplicate_columns, keep_strategy)
            
            # 5. 处理固定单元格数据提取
            if cleaning_options.get('fixed_cells_rules') and file_manager:
                result_df = self.extract_fixed_cells_data(
                    result_df, 
                    cleaning_options.get('fixed_cells_rules'), 
                    file_manager, 
                    session_id,
                    file_configs
                )
            
            # 更新最终统计
            self.processing_stats['output_rows'] = len(result_df)
            
            logging.info(f"数据处理完成：输入 {self.processing_stats['input_rows']} 行，"
                        f"输出 {self.processing_stats['output_rows']} 行，"
                        f"去重 {self.processing_stats['duplicates_removed']} 行，"
                        f"错误 {self.processing_stats['errors_count']} 个")
            
            return result_df
            
        except Exception as e:
            logging.error(f"数据处理时出错: {e}")
            raise
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """获取处理摘要"""
        return {
            'stats': self.processing_stats.copy(),
            'errors': self.errors[:self.max_error_samples],
            'total_errors': len(self.errors)
        }
    
    def detect_column_types(self, df: pd.DataFrame, sample_size: int = 1000) -> Dict[str, str]:
        """
        检测列的数据类型
        
        Args:
            df: 输入DataFrame
            sample_size: 采样大小
            
        Returns:
            列类型字典
        """
        column_types = {}
        
        for col in df.columns:
            sample_data = df[col].dropna().head(sample_size)
            
            if len(sample_data) == 0:
                column_types[col] = 'unknown'
                continue
            
            # 检查是否为数值类型
            if sample_data.dtype in [np.int64, np.float64]:
                column_types[col] = 'numeric'
            elif self._contains_numeric_pattern(sample_data.astype(str)):
                column_types[col] = 'numeric'
            elif self._contains_date_pattern(sample_data):
                column_types[col] = 'date'
            else:
                column_types[col] = 'text'
        
        return column_types
    
    def apply_column_configuration(self, df: pd.DataFrame, cleaning_options: Dict[str, Any]) -> pd.DataFrame:
        """
        应用列配置（顺序、显示、重命名）
        
        Args:
            df: 输入DataFrame
            cleaning_options: 清洗选项，包含列配置
            
        Returns:
            配置后的DataFrame
        """
        try:
            result_df = df.copy()
            
            # 获取列配置
            column_order = cleaning_options.get('column_order', [])
            column_names = cleaning_options.get('column_names', {})
            hidden_columns = set(cleaning_options.get('hidden_columns', []))
            
            # 应用列顺序
            if column_order:
                # 只保留存在的列，并按指定顺序排列
                existing_columns = [col for col in column_order if col in result_df.columns]
                # 添加不在顺序列表中但存在的列
                remaining_columns = [col for col in result_df.columns if col not in existing_columns]
                final_column_order = existing_columns + remaining_columns
                result_df = result_df[final_column_order]
            
            # 隐藏指定列
            if hidden_columns:
                visible_columns = [col for col in result_df.columns if col not in hidden_columns]
                result_df = result_df[visible_columns]
            
            # 重命名列
            if column_names:
                # 只重命名存在的列
                rename_dict = {old_name: new_name for old_name, new_name in column_names.items() 
                             if old_name in result_df.columns}
                if rename_dict:
                    result_df = result_df.rename(columns=rename_dict)
            
            logging.info(f"应用列配置后：{len(result_df.columns)} 列，顺序: {list(result_df.columns)}")
            
            return result_df
            
        except Exception as e:
            logging.error(f"应用列配置时出错: {e}")
            return df  # 返回原始数据框
    
    def extract_fixed_cells_data(self, result_df: pd.DataFrame, fixed_cells_rules: List[Dict], 
                                file_manager, session_id: str = None, 
                                file_configs: List[Dict] = None) -> pd.DataFrame:
        """
        从原始文件中提取固定单元格数据并添加为新列
        
        Args:
            result_df: 已合并的DataFrame
            fixed_cells_rules: 固定单元格提取规则列表 (通用规则，不绑定特定文件)
            file_manager: 文件管理器实例，用于读取单元格数据
            session_id: 用户会话ID
            file_configs: 文件配置列表，包含文件ID和数据源名称的映射
            
        Returns:
            添加了固定单元格列的DataFrame
        """
        try:
            if not fixed_cells_rules:
                return result_df
            
            logging.info(f"开始处理 {len(fixed_cells_rules)} 个固定单元格配置")
            logging.info(f"固定单元格规则: {fixed_cells_rules}")
            
            # 检查是否存在"数据源"列，用于匹配文件
            if '数据源' not in result_df.columns:
                logging.warning("结果DataFrame中没有'数据源'列，无法按文件匹配固定单元格")
                return result_df
            
            # 构建数据源名称到文件ID的映射
            source_to_file = {}
            if file_configs:
                for file_config in file_configs:
                    # 支持两种命名方式:驼峰(fileId/sourceName)和下划线(file_id/source_name)
                    source_name = file_config.get('source_name') or file_config.get('sourceName')
                    file_id = file_config.get('file_id') or file_config.get('fileId')
                    if source_name and file_id:
                        source_to_file[source_name] = file_id
                        logging.debug(f"映射: {source_name} -> {file_id}")
            
            logging.info(f"数据源到文件映射: {source_to_file}")
            
            if not source_to_file:
                logging.warning("没有有效的数据源到文件映射，无法提取固定单元格数据")
                return result_df
            
            # 按规则的列名分组，准备为每个新列创建数据
            columns_to_add = {}
            for rule in fixed_cells_rules:
                column_name = rule.get('column_name')
                if column_name and column_name not in columns_to_add:
                    columns_to_add[column_name] = {}  # 改为字典，key为数据源名，value为单元格值
            
            # 步骤1: 为每个文件读取一次固定单元格值（每个文件+规则只读取一次）
            file_cell_cache = {}  # 缓存: (file_id, sheet_name, cell_address) -> cell_value
            
            for data_source, file_id in source_to_file.items():
                for rule in fixed_cells_rules:
                    try:
                        column_name = rule.get('column_name')
                        cell_address = rule.get('cell_address')
                        sheet_name = rule.get('sheet_name')
                        
                        if not all([column_name, cell_address, sheet_name]):
                            continue
                        
                        cache_key = (file_id, sheet_name, cell_address)
                        
                        # 如果已经读取过，直接使用缓存
                        if cache_key not in file_cell_cache:
                            logging.debug(f"读取固定单元格: 数据源={data_source}, 文件={file_id}, 工作表={sheet_name}, 单元格={cell_address}")
                            
                            # 从该文件的指定工作表读取固定单元格的值
                            cell_value = file_manager.read_cell_value_by_address(
                                file_id=file_id,
                                sheet_name=sheet_name,
                                cell_address=cell_address,
                                session_id=session_id
                            )
                            
                            # 转换为字符串格式
                            cell_value = str(cell_value) if cell_value is not None and cell_value != "" else ""
                            file_cell_cache[cache_key] = cell_value
                            logging.debug(f"读取到的值: {cell_value}")
                        else:
                            cell_value = file_cell_cache[cache_key]
                        
                        # 存储: 该数据源对应该列的值
                        if column_name not in columns_to_add:
                            columns_to_add[column_name] = {}
                        columns_to_add[column_name][data_source] = cell_value
                        
                    except Exception as e:
                        logging.error(f"读取数据源'{data_source}'的固定单元格时出错: {e}", exc_info=True)
                        continue
            
            logging.info(f"已从 {len(source_to_file)} 个文件读取固定单元格值")
            
            # 步骤2: 批量填充到DataFrame（根据每行的数据源）
            for column_name, source_values in columns_to_add.items():
                # 创建新列，根据数据源填充值
                result_df[column_name] = result_df['数据源'].map(lambda x: source_values.get(x, ""))
                logging.info(f"成功添加固定单元格列: {column_name}，包含 {len(source_values)} 个不同数据源的值")
            
            logging.info(f"固定单元格处理完成: 共添加 {len(columns_to_add)} 列")
            
            if columns_to_add:
                logging.info(f"成功添加 {len(columns_to_add)} 个固定单元格列")
            
            return result_df
            
        except Exception as e:
            logging.error(f"提取固定单元格数据时出错: {e}")
            return result_df