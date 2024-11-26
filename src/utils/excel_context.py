"""Excel上下文管理器模块"""

import logging
from typing import Optional
import win32com.client
from contextlib import contextmanager
from src.core.exceptions import ExcelOperationError
from pathlib import Path

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
        # 使用GetActiveObject尝试获取已存在的Excel实例
        try:
            excel = win32com.client.GetActiveObject('Excel.Application')
            logger.debug("已获取现有Excel实例")
        except:
            # 如果没有现有实例，创建新的
            excel = win32com.client.DispatchEx('Excel.Application')
            logger.debug("已创建新的Excel实例")
        
        excel.DisplayAlerts = display_alerts
        excel.Visible = visible
        yield excel
    except Exception as e:
        raise ExcelOperationError(f"Excel操作失败: {str(e)}")
    finally:
        if excel:
            try:
                # 关闭所有打开的工作簿
                for wb in excel.Workbooks:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                excel.Quit()
                logger.debug("Excel应用程序已关闭")
            except:
                logger.warning("关闭Excel应用程序时出错", exc_info=True)

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
        workbook = excel_app.Workbooks.Open(
            abs_path,
            ReadOnly=read_only,
            UpdateLinks=False,  # 不更新链接
            IgnoreReadOnlyRecommended=True  # 忽略只读推荐
        )
        logger.debug(f"已打开工作簿: {abs_path}")
        yield workbook
    except Exception as e:
        raise ExcelOperationError(f"打开工作簿失败: {str(e)}")
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
                logger.debug(f"已关闭工作簿: {file_path}")
            except:
                logger.warning("关闭工作簿时出错", exc_info=True)
