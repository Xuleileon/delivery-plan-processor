import pytest
from pathlib import Path
from main import process_delivery_plan

def test_process_delivery_plan_with_optimization_file():
    """测试处理表优化.xlsx文件的完整流程"""
    # 设置输入文件路径
    input_file = str(Path(__file__).parent / "表优化.xlsx")
    
    # 执行处理流程
    result = process_delivery_plan(input_file)
    
    # 验证处理结果
    assert result['success'], f"处理失败: {result['message']}"
    assert 'data' in result, "返回结果中缺少data字段"
    assert all(key in result['data'] for key in ['preprocessed_file', 'transformed_file', 'final_file']), \
        "返回结果中缺少必要的文件路径"
    
    # 验证生成的文件是否存在
    for file_path in result['data'].values():
        assert Path(file_path).exists(), f"输出文件不存在: {file_path}"
        assert Path(file_path).stat().st_size > 0, f"输出文件为空: {file_path}"
    
    print("\n处理完成！")
    print(f"预处理文件: {result['data']['preprocessed_file']}")
    print(f"转换后文件: {result['data']['transformed_file']}")
    print(f"最终文件: {result['data']['final_file']}")
