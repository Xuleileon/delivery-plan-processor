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
            self.logger.error(f"预处理失败: {str(e)}", exc_info=True)
            if isinstance(e, (FileOperationError, ExcelOperationError)):
                raise
            raise ExcelOperationError(f"预处理失败: {str(e)}")
    
    def _process_sheets(self, workbook) -> None:
        """
        处理工作表
        
        Args:
            workbook: 工作簿对象
            
        Raises:
            ExcelOperationError: 工作表处理失败时抛出
        """
        sheets_to_process = config.excel_config['sheets_to_keep']
        processed_sheets = []
        
        # 验证工作簿是否包含工作表
        try:
            if not workbook.Worksheets.Count:
                raise ExcelOperationError("工作簿中没有工作表")
        except Exception as e:
            raise ExcelOperationError(f"无法访问工作表: {str(e)}")
            
        # 处理所有可用的工作表
        for sheet_name in sheets_to_process:
            try:
                # 使用Worksheets而不是Sheets
                sheet = workbook.Worksheets(sheet_name)
                
                # 确保所有公式都被计算
                try:
                    sheet.Calculate()
                except:
                    self.logger.warning(f"工作表 {sheet_name} 的公式计算失败")
                
                # 获取使用范围并复制值
                try:
                    used_range = sheet.UsedRange
                    if used_range:
                        used_range.Copy()
                        used_range.PasteSpecial(Paste=-4163)  # xlPasteValues
                        workbook.Application.CutCopyMode = False
                    else:
                        self.logger.warning(f"工作表 {sheet_name} 没有数据")
                except Exception as e:
                    raise ExcelOperationError(f"复制工作表 {sheet_name} 数据失败: {str(e)}")
                
                processed_sheets.append(sheet_name)
                self.logger.info(f"已处理工作表: {sheet_name}")
                
            except Exception as e:
                self.logger.error(f"处理工作表 {sheet_name} 时出错: {str(e)}")
                # 如果是关键工作表，则抛出异常
                if sheet_name in config.excel_config.get('required_sheets', []):
                    raise ExcelOperationError(f"处理必需的工作表 {sheet_name} 失败: {str(e)}")
        
        # 如果没有成功处理任何工作表，尝试处理第一个工作表
        if not processed_sheets:
            try:
                first_sheet = workbook.Worksheets(1)
                sheet_name = first_sheet.Name
                
                try:
                    first_sheet.Calculate()
                except:
                    self.logger.warning(f"工作表 {sheet_name} 的公式计算失败")
                
                used_range = first_sheet.UsedRange
                if used_range:
                    used_range.Copy()
                    used_range.PasteSpecial(Paste=-4163)
                    workbook.Application.CutCopyMode = False
                else:
                    raise ExcelOperationError(f"工作表 {sheet_name} 没有数据")
                
                self.logger.warning(f"未找到配置的工作表，已处理第一个工作表: {sheet_name}")
                processed_sheets.append(sheet_name)
                
            except Exception as e:
                self.logger.error(f"处理第一个工作表时出错: {str(e)}")
                raise ExcelOperationError(f"无法处理任何工作表: {str(e)}")
    
    def _remove_unused_sheets(self, workbook) -> None:
        """
        删除不需要的工作表
        
        Args:
            workbook: 工作簿对象
        """
        sheets_to_keep = config.excel_config['sheets_to_keep']
        
        # 获取所有工作表名称
        sheet_names = [sheet.Name for sheet in workbook.Sheets]
        
        # 如果没有找到任何需要保留的工作表，保留第一个工作表
        if not any(name in sheets_to_keep for name in sheet_names):
            self.logger.warning("未找到配置中指定的工作表，将保留第一个工作表")
            sheets_to_keep = [sheet_names[0]]
        
        # 从后往前删除工作表，避免索引变化问题
        for sheet_name in reversed(sheet_names):
            try:
                if sheet_name not in sheets_to_keep:
                    sheet = workbook.Sheets(sheet_name)
                    sheet.Delete()
                    self.logger.debug(f"已删除工作表: {sheet_name}")
            except Exception as e:
                self.logger.warning(f"删除工作表 {sheet_name} 时出错: {str(e)}")
                # 继续处理其他工作表，不抛出异常