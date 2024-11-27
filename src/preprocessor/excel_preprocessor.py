import logging
from pathlib import Path
from typing import List
from src.core.config import config
from src.core.exceptions import FileOperationError, ExcelOperationError
from src.utils.excel_context import excel_context, workbook_context
from src.utils.excel_utils import generate_output_path

class ExcelPreprocessor:
    """Excel预处理器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def process(self, input_file: Path) -> Path:
        """
        预处理Excel文件
        
        Args:
            input_file: 输入文件路径
            
        Returns:
            处理后的文件路径
            
        Raises:
            FileOperationError: 文件操作失败时抛出
            ExcelOperationError: Excel操作失败时抛出
        """
        if not input_file.exists():
            raise FileOperationError(f"输入文件不存在: {input_file}")
            
        try:
            with excel_context() as excel:
                with workbook_context(excel, str(input_file.absolute())) as workbook:
                    # 1. 处理工作表
                    self._process_sheets(workbook)
                    
                    # 2. 删除不需要的工作表
                    self._remove_unused_sheets(workbook)
                    
                    # 3. 保存结果
                    output_path = generate_output_path(
                        input_file,
                        Path(config.output_config['directory']),
                        config.output_config['suffixes']['preprocess']
                    )
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    workbook.SaveAs(str(output_path.absolute()))
                    
            self.logger.info(f"预处理完成，文件已保存至: {output_path}")
            return output_path
            
        except Exception as e:
            raise ExcelOperationError(f"预处理Excel文件失败: {str(e)}")
            
    def _process_sheets(self, workbook) -> None:
        """处理工作表"""
        for sheet in workbook.Sheets:
            sheet_name = sheet.Name
            if sheet_name in config.excel_config['sheets_to_keep']:
                self.logger.info(f"处理工作表: {sheet_name}")
                
                # 1. 清理数据
                self._clean_sheet_data(sheet)
                
                # 2. 处理公式
                self._process_formulas(sheet)
                
                # 3. 处理格式
                self._process_formats(sheet)
                
    def _clean_sheet_data(self, sheet) -> None:
        """清理工作表数据"""
        # 获取已使用区域
        used_range = sheet.UsedRange
        
        # 删除空行和空列
        self._delete_empty_rows(used_range)
        self._delete_empty_columns(used_range)
        
    def _process_formulas(self, sheet) -> None:
        """处理工作表中的公式"""
        # 获取已使用区域
        used_range = sheet.UsedRange
        
        # 将公式转换为值
        used_range.Value = used_range.Value
        
    def _process_formats(self, sheet) -> None:
        """处理工作表格式"""
        # 获取已使用区域
        used_range = sheet.UsedRange
        
        # 设置表头格式
        header_row = used_range.Rows(1)
        header_row.Font.Bold = True
        header_row.Interior.Color = int(config.styles_config['header']['green'], 16)
        
    def _delete_empty_rows(self, used_range) -> None:
        """删除空行"""
        rows = used_range.Rows
        for i in range(rows.Count, 0, -1):
            row = rows.Item(i)
            if not any(cell.Value for cell in row.Cells):
                row.Delete()
                
    def _delete_empty_columns(self, used_range) -> None:
        """删除空列"""
        columns = used_range.Columns
        for i in range(columns.Count, 0, -1):
            column = columns.Item(i)
            if not any(cell.Value for cell in column.Cells):
                column.Delete()
                
    def _remove_unused_sheets(self, workbook) -> None:
        """删除不需要的工作表"""
        sheets_to_keep = config.excel_config['sheets_to_keep']
        
        for sheet in workbook.Sheets:
            if sheet.Name not in sheets_to_keep:
                self.logger.info(f"删除工作表: {sheet.Name}")
                sheet.Delete()