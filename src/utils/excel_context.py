"""Excel上下文管理器模块"""

import logging
from typing import Optional
import win32com.client
from contextlib import contextmanager
from src.core.exceptions import ExcelOperationError
from pathlib import Path
import pythoncom

logger = logging.getLogger(__name__)

@contextmanager
def excel_context(visible: bool = False, display_alerts: bool = False):
    """
    Excel应用程序上下文管理器
    
    Args:
        visible: 是否显示Excel窗口
        display_alerts: 是否显示Excel警告
    
    Yields:
        Excel.Application对象
    
    Raises:
        ExcelOperationError: Excel操作失败时抛出
    """
    excel = None
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        try:
            # 尝试使用DispatchEx创建新实例
            excel = win32com.client.DispatchEx('Excel.Application')
            logger.debug("已创建新的Excel实例")
            
            # 设置Excel应用程序属性
            excel.DisplayAlerts = display_alerts
            excel.Visible = visible
            excel.ScreenUpdating = False  # 禁用屏幕更新以提高性能
            
            yield excel
            
        except pythoncom.com_error as e:
            raise ExcelOperationError(f"创建Excel实例失败: {str(e)}")
            
    except Exception as e:
        if not isinstance(e, ExcelOperationError):
            raise ExcelOperationError(f"Excel操作失败: {str(e)}")
        raise
        
    finally:
        if excel:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                # 关闭所有打开的工作簿
                for wb in excel.Workbooks:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        logger.warning("关闭工作簿时出错")
                excel.Quit()
            except:
                logger.warning("关闭Excel应用程序时出错")
            finally:
                pythoncom.CoUninitialize()

@contextmanager
def workbook_context(excel_app, file_path: str, read_only: bool = False):
    """
    Excel工作簿上下文管理器
    
    Args:
        excel_app: Excel.Application对象
        file_path: 工作簿文件路径
        read_only: 是否以只读方式打开
    
    Yields:
        Workbook对象
    
    Raises:
        ExcelOperationError: Excel操作失败时抛出
    """
    workbook = None
    try:
        # 确保文件路径是绝对路径
        abs_path = str(Path(file_path).absolute())
        
        # 检查文件是否被锁定
        try:
            with open(abs_path, 'rb') as f:
                pass
        except IOError:
            raise ExcelOperationError(f"文件被锁定或无法访问: {abs_path}")
            
        try:
            workbook = excel_app.Workbooks.Open(
                abs_path,
                ReadOnly=read_only,
                UpdateLinks=False,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=2  # xlRepairFile
            )
            logger.debug(f"已打开工作簿: {abs_path}")
            
            # 验证工作簿是否正确打开
            if not workbook or not hasattr(workbook, 'Worksheets'):
                raise ExcelOperationError("工作簿打开失败或格式无效")
                
            yield workbook
            
        except pythoncom.com_error as e:
            raise ExcelOperationError(f"打开工作簿失败: {str(e)}")
            
    except Exception as e:
        if not isinstance(e, ExcelOperationError):
            raise ExcelOperationError(f"工作簿操作失败: {str(e)}")
        raise
        
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                logger.warning("关闭工作簿时出错")
