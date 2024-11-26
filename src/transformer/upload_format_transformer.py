import logging
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
from src.utils.excel_utils import generate_output_path

class UploadFormatTransformer:
    """将合并后的到货计划转换为上传格式"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def transform(self, input_file: Path) -> Path:
        """
        转换到货计划为上传格式
        
        Args:
            input_file: 输入文件路径（合并后的Excel文件）
            
        Returns:
            转换后的文件路径
        """
        try:
            # 读取合并后的数据
            df = pd.read_excel(input_file, sheet_name='汇总')
            
            # 获取所有日期列
            date_cols = [col for col in df.columns if pd.notna(pd.to_datetime(col, errors='coerce'))]
            date_cols.sort()  # 确保日期按顺序排列
            
            # 创建60天的日期列
            today = datetime.now()
            day_cols = [f'day{i+1}' for i in range(60)]
            
            # 创建结果DataFrame
            result_rows = []
            
            # 处理每个SKU
            for _, row in df.iterrows():
                sku = row['sku编码']
                if pd.isna(sku):
                    continue
                    
                # 从规格中提取颜色和尺码
                spec = str(row['规格']) if pd.notna(row['规格']) else ''
                color = ''
                size = ''
                if '/' in spec:
                    parts = spec.split('/')
                    if len(parts) >= 2:
                        color = parts[0].strip()
                        size = parts[1].strip()
                
                # 创建新行
                result_row = {
                    'sku_no': sku,
                    'color': color,
                    'size': size
                }
                
                # 初始化所有天数为0
                for day_col in day_cols:
                    result_row[day_col] = 0
                    
                # 填充实际数据
                for date in date_cols:
                    qty = row[date]
                    if pd.notna(qty) and qty > 0:
                        # 计算这个日期是第几天
                        date_diff = (pd.to_datetime(date) - today).days
                        if 0 <= date_diff < 60:  # 只处理60天内的数据
                            day_col = f'day{date_diff + 1}'
                            result_row[day_col] = int(qty)
                
                # 添加dt字段（当前时间戳）
                result_row['dt'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                result_rows.append(result_row)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(result_rows)
            
            # 确保所有列都存在并按正确顺序排列
            all_cols = ['sku_no', 'color', 'size'] + day_cols + ['dt']
            result_df = result_df.reindex(columns=all_cols)
            
            # 保存结果
            output_path = generate_output_path(input_file, Path("output"), "上传格式")
            result_df.to_excel(output_path, index=False, sheet_name='Sheet1')
            
            self._log_results(df, result_df)
            return output_path
            
        except Exception as e:
            self.logger.error(f"转换失败: {str(e)}", exc_info=True)
            raise
            
    def _log_results(self, original_df, result_df):
        """记录处理结果"""
        self.logger.info(f"原始SKU数: {len(original_df)}")
        self.logger.info(f"转换后SKU数: {len(result_df)}")
        non_zero_days = len([col for col in result_df.columns if col.startswith('day') and (result_df[col] > 0).any()])
        self.logger.info(f"包含数据的天数: {non_zero_days}")
