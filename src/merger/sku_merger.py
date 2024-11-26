import logging
from pathlib import Path
import pandas as pd
from src.utils.excel_utils import generate_output_path, format_sku

class SKUMerger:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def merge(self, input_file: Path) -> Path:
        """合并SKU数据"""
        try:
            # 读取数据，不指定dtype，让pandas自动推断类型
            df = pd.read_excel(input_file, sheet_name='汇总')
            
            # 处理数据
            df = self._format_data(df)
            duplicates = self._find_duplicates(df)
            merged_df = self._merge_duplicates(df)
            
            # 保存结果
            output_path = generate_output_path(input_file, Path("output"), "合并")
            self._save_result(merged_df, output_path)
            
            self._log_results(df, merged_df, duplicates)
            return output_path
            
        except Exception as e:
            self.logger.error(f"合并失败: {str(e)}", exc_info=True)
            raise
            
    def _format_data(self, df):
        """格式化数据"""
        # 只对SKU相关列应用format_sku
        sku_cols = ['sku_no']
        for col in sku_cols:
            if col in df.columns:
                df[col] = df[col].apply(format_sku)
        return df
        
    def _find_duplicates(self, df):
        """找出重复的SKU"""
        return df[df.duplicated(subset=['sku_no'], keep=False)]
        
    def _merge_duplicates(self, df):
        """合并重复的SKU"""
        # 识别需要合并的列
        date_cols = [col for col in df.columns if col.startswith('day')]
        sum_cols = date_cols
        first_cols = ['sku_no', 'color', 'size']
        
        # 构建聚合字典
        agg_dict = {}
        for col in df.columns:
            if col in sum_cols:
                agg_dict[col] = 'sum'
            elif col in first_cols:
                agg_dict[col] = 'first'
            elif col == 'dt':
                agg_dict[col] = 'first'
        
        # 按SKU分组并合并
        merged_df = df.groupby('sku_no', as_index=False).agg(agg_dict)
        
        # 确保数值列为float类型
        for col in date_cols:
            merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce').fillna(0.0)
            
        return merged_df
        
    def _save_result(self, df, output_path):
        """保存结果"""
        df.to_excel(output_path, index=False, sheet_name='汇总')
        
    def _log_results(self, original_df, merged_df, duplicates):
        """记录处理结果"""
        self.logger.info(f"原始记录数: {len(original_df)}")
        self.logger.info(f"合并后记录数: {len(merged_df)}")
        self.logger.info(f"重复SKU数: {len(duplicates)}")