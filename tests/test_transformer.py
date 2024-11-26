"""数据转换器测试模块"""

import pytest
import pandas as pd
from pathlib import Path
from src.transformer.delivery_plan_transformer import DeliveryPlanTransformer
from src.core.exceptions import FileOperationError, DataTransformError

def test_transform_success(test_config, sample_excel_file):
    """测试正常转换流程"""
    transformer = DeliveryPlanTransformer()
    output_path = transformer.transform(Path(sample_excel_file))
    
    assert output_path.exists()
    
    # 验证输出文件内容
    df = pd.read_excel(output_path, sheet_name='汇总')
    assert len(df) == 4  # 2个常规产品 + 2个S级产品
    assert all(col in df.columns for col in ['编码', '名称', '规格', '数量', '日期'])
    
    # 验证日期格式
    assert all(pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d') == df['日期'])

def test_transform_file_not_found(test_config, test_data_dir):
    """测试文件不存在的情况"""
    transformer = DeliveryPlanTransformer()
    non_existent_file = test_data_dir / 'non_existent.xlsx'
    
    with pytest.raises(FileOperationError):
        transformer.transform(non_existent_file)

def test_transform_invalid_sheet(test_config, test_data_dir):
    """测试工作表格式错误的情况"""
    # 创建一个缺少必要工作表的Excel文件
    file_path = test_data_dir / 'invalid.xlsx'
    df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
    df.to_excel(file_path, sheet_name='Sheet1', index=False)
    
    transformer = DeliveryPlanTransformer()
    with pytest.raises(DataTransformError):
        transformer.transform(file_path)

def test_transform_invalid_date(test_config, test_data_dir):
    """测试日期格式错误的情况"""
    # 创建包含无效日期的Excel文件
    file_path = test_data_dir / 'invalid_date.xlsx'
    
    regular_data = pd.DataFrame({
        '编码': ['R001'],
        '名称': ['常规产品1'],
        '规格': ['规格1'],
        '数量': [100],
        '日期': ['invalid_date']  # 无效日期
    })
    
    s_level_data = pd.DataFrame({
        '编码': ['S001'],
        '名称': ['S级产品1'],
        '规格': ['规格1'],
        '数量': [200],
        '日期': ['2023-01-01']
    })
    
    with pd.ExcelWriter(file_path) as writer:
        regular_data.to_excel(writer, sheet_name='常规产品', index=False)
        s_level_data.to_excel(writer, sheet_name='S级产品', index=False)
        pd.DataFrame().to_excel(writer, sheet_name='汇总', index=False)
    
    transformer = DeliveryPlanTransformer()
    with pytest.raises(DataTransformError):
        transformer.transform(file_path)
