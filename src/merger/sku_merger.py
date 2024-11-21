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
            output_path = generate_output_path(input_file, "合并")
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
        
    # ... 其他具体实现方法 ... 