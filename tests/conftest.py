"""测试配置模块"""

import os
import shutil
import pytest
from pathlib import Path
import pandas as pd
from src.core.config import ConfigManager

@pytest.fixture(scope="session")
def test_data_dir(tmp_path_factory):
    """创建测试数据目录"""
    test_dir = tmp_path_factory.mktemp("test_data")
    yield test_dir
    # 清理测试数据
    shutil.rmtree(test_dir)

@pytest.fixture(scope="session")
def test_config(test_data_dir):
    """创建测试配置"""
    config = {
        'excel': {
            'sheets': {
                'process': [
                    {'name': '常规产品', 'type': 'regular'},
                    {'name': 'S级产品', 'type': 's_level'},
                    {'name': '汇总', 'type': 'summary'}
                ],
                'keep': ['常规产品', 'S级产品', '汇总']
            },
            'header_styles': {
                'basic_columns': ['编码', '名称', '规格', '数量'],
                'colors': {
                    'basic': '1F6B3B',
                    'other': 'F4B183'
                }
            }
        },
        'output': {
            'directory': str(test_data_dir / 'output'),
            'suffixes': {
                'preprocess': '预处理',
                'transform': '转换',
                'merge': '合并'
            }
        },
        'date': {
            'input_formats': ['%Y-%m-%d', '%Y/%m/%d'],
            'output_format': '%Y-%m-%d'
        }
    }
    
    config_path = test_data_dir / 'test_settings.yaml'
    os.makedirs(config['output']['directory'], exist_ok=True)
    
    # 重置配置管理器
    ConfigManager._instance = None
    ConfigManager._config = config
    
    return config

@pytest.fixture
def sample_excel_file(test_data_dir):
    """创建示例Excel文件"""
    file_path = test_data_dir / 'test.xlsx'
    
    # 创建示例数据
    regular_data = pd.DataFrame({
        '编码': ['R001', 'R002'],
        '名称': ['常规产品1', '常规产品2'],
        '规格': ['规格1', '规格2'],
        '数量': [100, 200],
        '日期': ['2023-01-01', '2023-01-02']
    })
    
    s_level_data = pd.DataFrame({
        '编码': ['S001', 'S002'],
        '名称': ['S级产品1', 'S级产品2'],
        '规格': ['规格1', '规格2'],
        '数量': [300, 400],
        '日期': ['2023-01-03', '2023-01-04']
    })
    
    summary_data = pd.DataFrame()
    
    # 创建Excel文件
    with pd.ExcelWriter(file_path) as writer:
        regular_data.to_excel(writer, sheet_name='常规产品', index=False)
        s_level_data.to_excel(writer, sheet_name='S级产品', index=False)
        summary_data.to_excel(writer, sheet_name='汇总', index=False)
    
    return file_path
