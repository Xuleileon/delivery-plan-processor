import logging
from pathlib import Path
import win32com.client
from src.utils.excel_utils import generate_output_path

class ExcelPreprocessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def process(self, input_file: Path) -> Path:
        """预处理Excel文件"""
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.DisplayAlerts = False
            
            self.logger.info(f"开始处理文件: {input_file}")
            workbook = excel.Workbooks.Open(str(input_file.absolute()))
            
            # 处理工作表
            sheets_to_calculate = ['常规产品', 'S级产品']
            sheets_to_keep = ['常规产品', 'S级产品', '汇总']
            
            self._calculate_sheets(workbook, sheets_to_calculate)
            self._clean_sheets(workbook, sheets_to_keep)
            
            # 保存结果
            output_path = generate_output_path(input_file, "预处理")
            workbook.SaveAs(str(output_path))
            workbook.Close()
            excel.Quit()
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"预处理失败: {str(e)}", exc_info=True)
            raise
            
    def _calculate_sheets(self, workbook, sheet_names):
        """计算指定工作表的公式"""
        for sheet_name in sheet_names:
            sheet = workbook.Sheets(sheet_name)
            sheet.Calculate()
            used_range = sheet.UsedRange
            used_range.Copy()
            used_range.PasteSpecial(Paste=-4163)
            
    def _clean_sheets(self, workbook, sheets_to_keep):
        """清理多余的工作表"""
        for i in range(workbook.Sheets.Count, 0, -1):
            sheet = workbook.Sheets(i)
            if sheet.Name not in sheets_to_keep:
                sheet.Delete() 