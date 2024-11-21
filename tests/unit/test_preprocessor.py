import pytest
from pathlib import Path
from src.preprocessor.excel_preprocessor import ExcelPreprocessor
from src.utils.excel_utils import setup_logging

class TestExcelPreprocessor:
    @pytest.fixture(scope="class")
    def setup_test_env(self):
        # 创建日志目录
        Path("logs").mkdir(exist_ok=True)
        # 创建测试数据目录
        Path("tests/fixtures").mkdir(parents=True, exist_ok=True)
        Path("output").mkdir(exist_ok=True)
        # 设置日志
        setup_logging()
        
    @pytest.fixture
    def preprocessor(self, setup_test_env):
        return ExcelPreprocessor()
    
    @pytest.fixture
    def sample_excel(self, tmp_path):
        """创建一个测试用的Excel文件"""
        import openpyxl
        wb = openpyxl.Workbook()
        
        # 创建测试工作表
        for sheet_name in ['常规产品', 'S级产品', '汇总']:
            ws = wb.create_sheet(sheet_name)
            ws['A1'] = 'SKU编码'
            ws['B1'] = '数量'
        
        # 删除默认的Sheet
        wb.remove(wb['Sheet'])
        
        # 保存测试文件
        test_file = tmp_path / "sample_excel.xlsx"
        wb.save(test_file)
        return test_file
    
    def test_process_valid_file(self, preprocessor, sample_excel):
        """测试处理有效文件"""
        result = preprocessor.process(sample_excel)
        assert result.exists()
        assert result.suffix == '.xlsx'
    
    def test_process_invalid_file(self, preprocessor):
        """测试处理不存在的文件"""
        with pytest.raises(FileNotFoundError, match="文件不存在"):
            preprocessor.process(Path("not_exists.xlsx"))
    
    @pytest.mark.slow
    def test_large_file_processing(self, preprocessor, tmp_path):
        """测试大文件处理"""
        # TODO: 实现大文件测试
        pass
    
    @pytest.mark.integration
    def test_full_workflow(self, preprocessor, sample_excel):
        """测试完整工作流"""
        # 处理文件
        result = preprocessor.process(sample_excel)
        
        # 验证输出
        assert result.exists()
        
        # 验证输出文件包含所需的工作表
        import openpyxl
        wb = openpyxl.load_workbook(result)
        assert set(wb.sheetnames) == {'常规产品', 'S级产品', '汇总'}