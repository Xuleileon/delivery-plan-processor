import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from copy import copy
import sys

def setup_logging():
    """设置日志配置"""
    # 获取应用程序的基础路径
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        base_path = Path(sys._MEIPASS)
    else:
        # 如果是开发环境
        base_path = Path.cwd()
    
    # 创建logs目录
    log_dir = base_path / 'logs'
    log_dir.mkdir(exist_ok=True)
    
    # 设置日志文件路径
    log_file = log_dir / f'excel_processor_{datetime.now():%Y%m%d}.log'
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

def generate_output_path(input_file: Path, output_dir: Path, suffix: str) -> Path:
    """
    生成输出文件路径
    
    Args:
        input_file: 输入文件路径
        output_dir: 输出目录路径
        suffix: 文件后缀标识
    
    Returns:
        Path: 输出文件路径
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return output_dir / f"{input_file.stem}_{suffix}_{timestamp}.xlsx"

def format_sku(x):
    """格式化SKU编码"""
    if pd.isna(x):
        return ''
    if isinstance(x, (int, float)):
        return str(int(x))
    return str(x).strip()

def copy_cell_format(source_cell, target_cell):
    """复制单元格格式"""
    target_cell.font = copy(source_cell.font)
    target_cell.border = copy(source_cell.border)
    target_cell.fill = copy(source_cell.fill)
    target_cell.alignment = copy(source_cell.alignment)