import logging
from pathlib import Path
import pandas as pd
from src.utils.excel_utils import generate_output_path, copy_cell_format
import openpyxl
from openpyxl.styles import PatternFill

class DeliveryPlanTransformer:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def transform(self, input_file: Path) -> Path:
        """转换送货计划"""
        try:
            # 读取数据
            sheets = self._read_sheets(input_file)
            
            # 转换格式
            transformed_df = self._transform_format(sheets)
            
            # 保存结果
            output_path = generate_output_path(input_file, Path("output"), "转换")
            self._save_result(transformed_df, output_path, input_file)
            
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
        
    def _transform_format(self, sheets):
        """转换数据格式"""
        # 合并常规产品和S级产品
        df = pd.concat([sheets['regular'], sheets['s_level']], ignore_index=True)
        
        # 转换日期格式
        date_cols = df.columns[df.columns.str.contains('日期|时间')]
        for col in date_cols:
            df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
        return df
        
    def _save_result(self, df, output_path, input_file):
        """保存结果并设置格式"""
        # 保存数据
        df.to_excel(output_path, index=False, sheet_name='汇总')
        
        # 设置格式
        wb = openpyxl.load_workbook(output_path)
        ws = wb['汇总']
        
        # 复制原始文件的格式
        source_wb = openpyxl.load_workbook(input_file)
        source_ws = source_wb['汇总']
        
        # 设置表头颜色
        green_fill = PatternFill(start_color='1F6B3B', end_color='1F6B3B', fill_type='solid')
        orange_fill = PatternFill(start_color='F4B183', end_color='F4B183', fill_type='solid')
        
        for col in ws.iter_cols(1, ws.max_column):
            cell = col[0]  # 表头单元格
            if any(keyword in cell.value for keyword in ['编码', '名称', '规格', '数量']):
                cell.fill = green_fill
            else:
                cell.fill = orange_fill
                
        wb.save(output_path)