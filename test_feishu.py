"""测试从飞书获取数据并处理"""
import os
from pathlib import Path
from main import process_delivery_plan

def test_feishu_process():
    """测试从飞书获取数据并处理"""
    try:
        # 设置飞书应用凭证
        os.environ['FEISHU_APP_ID'] = 'cli_a63afccc31bc900b'
        os.environ['FEISHU_APP_SECRET'] = 'hJLJHYk64H6nCSz3aq77ThJnJUzkOAC5'
        
        print("开始从飞书获取数据...")
        # 从飞书获取数据并处理
        result = process_delivery_plan(
            input_source=None,  # None表示从飞书获取数据
            output_dir='output'
        )
        
        # 检查结果
        if result['success']:
            print("\n处理成功!")
            print(f"预处理文件: {result['data']['preprocessed_file']}")
            print(f"转换后文件: {result['data']['transformed_file']}")
            print(f"最终文件: {result['data']['final_file']}")
        else:
            print(f"\n处理失败: {result['message']}")
    except Exception as e:
        print(f"\n发生错误: {str(e)}")
        print("\n请确保:")
        print("1. 飞书应用有访问表格的权限")
        print("2. config.yaml 中的飞书表格配置正确")
        print("3. 表格中至少有一个可见的工作表")
        raise

if __name__ == "__main__":
    test_feishu_process()
