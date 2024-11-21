import logging
from pathlib import Path
import pandas as pd
from src.utils.excel_utils import generate_output_path

class DeliveryPlanTransformer:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def transform(self, input_file: Path) -> Path:
        """转换送货计划"""
        try:
            # 读取数据
            sheets = self._read_sheets(input_file)
            
            # 处理数据
            all_dates = self._collect_dates(sheets)
            transformed_data = self._transform_data(sheets, all_dates)
            
            # 保存结果
            output_path = generate_output_path(input_file, "转换")
            self._save_result(transformed_data, output_path)
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"转换失败: {str(e)}", exc_info=True)
            raise
            
    def _read_sheets(self, input_file):
        """读取所有工作表"""
        return {
            'regular': pd.read_excel(input_file, sheet_name='常规产品'),
            's_level': pd.read_excel(input_file, sheet_name='S级产品'),
            'summary': pd.read_excel(input_file, sheet_name='汇总')
        }
        
    # ... 其他具体实现方法 ... 