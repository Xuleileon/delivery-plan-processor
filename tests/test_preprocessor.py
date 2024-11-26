"""Excel预处理器测试模块"""

import pytest
import pandas as pd
from pathlib import Path
from src.preprocessor.excel_preprocessor import ExcelPreprocessor
from src.core.exceptions import FileOperationError, ExcelOperationError

def test_process_success(test_config, sample_excel_file):
    """测试正常预处理流程"""
    preprocessor = ExcelPreprocessor()
    output_path = preprocessor.process(Path(sample_excel_file))
    
    assert output_path.exists()
    
    # 验证输出文件包含所有必要的工作表
    with pd.ExcelFile(output_path) as xls:
        sheets = xls.sheet_names
        required_sheets = test_config['excel']['sheets']['keep']
        assert all(sheet in sheets for sheet in required_sheets)
        assert len(sheets) == len(required_sheets)  # 确保没有多余的工作表

def test_process_file_not_found(test_config, test_data_dir):
    """测试文件不存在的情况"""
    preprocessor = ExcelPreprocessor()
    non_existent_file = test_data_dir / 'non_existent.xlsx'
    
    with pytest.raises(FileOperationError):
        preprocessor.process(non_existent_file)

def test_process_invalid_file(test_config, test_data_dir):
    """测试无效Excel文件的情况"""
    # 创建一个无效的Excel文件
    invalid_file = test_data_dir / 'invalid.txt'
    invalid_file.write_text('This is not an Excel file')
    
    preprocessor = ExcelPreprocessor()
    with pytest.raises(ExcelOperationError):
        preprocessor.process(invalid_file)

def test_process_missing_sheets(test_config, test_data_dir):
    """测试缺少必要工作表的情况"""
    # 创建一个缺少必要工作表的Excel文件
    file_path = test_data_dir / 'missing_sheets.xlsx'
    df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
    df.to_excel(file_path, sheet_name='Sheet1', index=False)
    
    preprocessor = ExcelPreprocessor()
    with pytest.raises(ExcelOperationError):
        preprocessor.process(file_path)

def test_process_locked_file(test_config, test_data_dir):
    """测试文件被锁定的情况"""
    # 创建一个Excel文件
    file_path = test_data_dir / 'locked.xlsx'
    df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
    df.to_excel(file_path, sheet_name='Sheet1', index=False)
    
    # 模拟文件被锁定（通过打开文件但不关闭）
    f = open(file_path, 'rb')
    try:
        preprocessor = ExcelPreprocessor()
        with pytest.raises(ExcelOperationError):
            preprocessor.process(file_path)
    finally:
        f.close()  # 确保文件被关闭

def test_process_optimization_file():
    """测试处理表优化.xlsx文件"""
    preprocessor = ExcelPreprocessor()
    
    # 使用表优化.xlsx文件
    input_file = Path(__file__).parent / "表优化.xlsx"
    assert input_file.exists(), f"测试文件不存在: {input_file}"
    
    output_path = preprocessor.process(input_file)
    print(f"\n处理完成！输出文件保存在: {output_path}")
    
    assert output_path.exists(), "输出文件未生成"
    assert output_path.stat().st_size > 0, "输出文件是空的"
