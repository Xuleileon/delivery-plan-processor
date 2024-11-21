import logging
from pathlib import Path
from typing import Dict, Union, Tuple
import yaml
from src.preprocessor.excel_preprocessor import ExcelPreprocessor
from src.transformer.delivery_plan_transformer import DeliveryPlanTransformer
from src.merger.sku_merger import SKUMerger
from src.utils.excel_utils import setup_logging

def process_delivery_plan(
    input_file_path: str,
    output_dir: str = None,
    config_path: str = None
) -> Dict[str, Union[bool, str, Dict[str, str]]]:
    """
    处理到货计划的主函数
    
    Args:
        input_file_path: 输入文件路径
        output_dir: 输出目录路径（可选）
        config_path: 配置文件路径（可选）
    
    Returns:
        Dict: {
            'success': bool,
            'message': str,
            'data': {
                'preprocessed_file': str,
                'transformed_file': str,
                'final_file': str
            }
        }
    """
    # 设置日志
    setup_logging()
    logger = logging.getLogger(__name__)
    
    try:
        # 加载配置
        if config_path:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
        else:
            config = {}
        
        # 处理输入路径
        input_file = Path(input_file_path)
        if not input_file.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_file_path}")
            
        # 设置输出目录
        output_dir = Path(output_dir) if output_dir else input_file.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 1. 预处理阶段
        preprocessor = ExcelPreprocessor(config.get('excel', {}))
        preprocessed_file = preprocessor.process(input_file, output_dir)
        logger.info(f"预处理完成: {preprocessed_file}")

        # 2. 转换阶段
        transformer = DeliveryPlanTransformer(config.get('transformer', {}))
        transformed_file = transformer.transform(preprocessed_file, output_dir)
        logger.info(f"转换完成: {transformed_file}")

        # 3. 合并阶段
        merger = SKUMerger(config.get('merger', {}))
        final_file = merger.merge(transformed_file, output_dir)
        logger.info(f"合并完成: {final_file}")

        result = {
            'success': True,
            'message': "处理成功",
            'data': {
                'preprocessed_file': str(preprocessed_file),
                'transformed_file': str(transformed_file),
                'final_file': str(final_file)
            }
        }
        
        logger.info("整个流程处理完成!")
        return result

    except Exception as e:
        logger.error(f"处理失败: {str(e)}", exc_info=True)
        return {
            'success': False,
            'message': f"处理失败: {str(e)}",
            'data': {}
        }

def lambda_handler(event: Dict, context: Any) -> Dict:
    """AWS Lambda处理函数"""
    input_file = event.get('input_file')
    output_dir = event.get('output_dir')
    config_path = event.get('config_path')
    
    return process_delivery_plan(input_file, output_dir, config_path)

def local_handler():
    """本地处理函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='处理到货计划Excel文件')
    parser.add_argument('input_file', help='输入文件路径')
    parser.add_argument('--output-dir', help='输出目录路径')
    parser.add_argument('--config', help='配置文件路径')
    
    args = parser.parse_args()
    result = process_delivery_plan(args.input_file, args.output_dir, args.config)
    print(result['message'])
    return result

if __name__ == "__main__":
    local_handler() 