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
            # 读取数据
            df = pd.read_excel(input_file, sheet_name='汇总', dtype=str)
            
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
        for col in df.columns:
            df[col] = df[col].apply(format_sku)
        return df
        
    def _find_duplicates(self, df):
        """查找重复的SKU"""
        return df[df.duplicated(subset=['sku编码'], keep=False)]
        
    def _merge_duplicates(self, df):
        """合并重复的SKU"""
        # 按SKU分组并合并数量
        numeric_cols = df.columns[df.columns.str.contains('数量|交货量')]
        agg_dict = {col: 'sum' for col in numeric_cols}
        
        # 保留其他列的第一个值
        for col in df.columns:
            if col not in numeric_cols:
                agg_dict[col] = 'first'
                
        return df.groupby('sku编码', as_index=False).agg(agg_dict)
        
    def _save_result(self, df, output_path):
        """保存结果"""
        df.to_excel(output_path, index=False, sheet_name='汇总')
        
    def _log_results(self, original_df, merged_df, duplicates):
        """记录处理结果"""
        self.logger.info(f"原始记录数: {len(original_df)}")
        self.logger.info(f"合并后记录数: {len(merged_df)}")
        self.logger.info(f"重复SKU数: {len(duplicates)}")