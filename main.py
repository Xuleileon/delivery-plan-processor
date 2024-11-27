import logging
from pathlib import Path
from typing import Dict, Union, Tuple, Any
import yaml
from src.preprocessor.excel_preprocessor import ExcelPreprocessor
from src.transformer.delivery_plan_transformer import DeliveryPlanTransformer
from src.merger.sku_merger import SKUMerger
from src.utils.excel_utils import setup_logging
from src.utils.feishu_utils import FeishuSheetDownloader

def process_delivery_plan(
    input_source: Union[str, Dict] = None,
    output_dir: str = None,
    config_path: str = None
) -> Dict[str, Union[bool, str, Dict[str, str]]]:
    """
    处理到货计划的主函数
    
    Args:
        input_source: 输入来源，可以是Excel文件路径(str)或飞书配置(Dict)
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
            config_path = Path(__file__).parent / 'config' / 'config.yaml'
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
        
        # 设置输出目录
        output_dir = Path(output_dir) if output_dir else Path(config['output']['directory'])
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 根据输入来源获取数据
        if isinstance(input_source, str):
            # 使用本地Excel文件
            input_file = Path(input_source)
            if not input_file.exists():
                raise FileNotFoundError(f"输入文件不存在: {input_source}")
            initial_file = input_file
        else:
            # 从飞书获取数据
            logger.info("从飞书获取数据...")
            downloader = FeishuSheetDownloader(
                app_id=config['feishu']['app_id'],
                app_secret=config['feishu']['app_secret']
            )
            
            # 构建sheet URLs
            sheet_urls = []
            spreadsheet_token = config['feishu']['spreadsheet_token']
            sheet_config = config['feishu']['sheets']
            url = f"https://fr1r3d1ckr.feishu.cn/sheets/{spreadsheet_token}"
            sheet_urls.append(url)
            
            # 下载并保存为Excel文件
            initial_file = Path(downloader.download_sheets(sheet_urls, sheet_config))
            logger.info(f"飞书数据已保存到: {initial_file}")
        
        # 1. 预处理阶段 - 处理Excel公式和格式
        preprocessor = ExcelPreprocessor()
        preprocessed_file = preprocessor.process(initial_file)
        logger.info(f"预处理完成: {preprocessed_file}")
        
        # 2. 转换阶段 - 转换数据格式
        transformer = DeliveryPlanTransformer()
        transformed_file = transformer.transform(preprocessed_file)
        logger.info(f"转换完成: {transformed_file}")
        
        # 3. 合并阶段 - 合并SKU数据
        merger = SKUMerger()
        final_file = merger.merge(transformed_file)
        logger.info(f"合并完成: {final_file}")
        
        return {
            'success': True,
            'message': '处理成功',
            'data': {
                'preprocessed_file': str(preprocessed_file),
                'transformed_file': str(transformed_file),
                'final_file': str(final_file)
            }
        }
        
    except Exception as e:
        logger.error(f"处理失败: {str(e)}", exc_info=True)
        return {
            'success': False,
            'message': str(e),
            'data': {}
        }

def lambda_handler(event: Dict, context: Any) -> Dict:
    """AWS Lambda处理函数"""
    input_source = event.get('input_source')
    output_dir = event.get('output_dir')
    config_path = event.get('config_path')
    
    return process_delivery_plan(input_source, output_dir, config_path)

def local_handler():
    """本地处理函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='处理到货计划')
    parser.add_argument('--input', help='输入文件路径或使用"feishu"从飞书获取数据')
    parser.add_argument('--output', help='输出目录路径')
    parser.add_argument('--config', help='配置文件路径')
    
    args = parser.parse_args()
    
    # 如果指定使用飞书，则传入None作为input_source
    input_source = None if args.input == 'feishu' else args.input
    
    result = process_delivery_plan(input_source, args.output, args.config)
    if result['success']:
        print("处理成功！")
        print(f"预处理文件: {result['data']['preprocessed_file']}")
        print(f"转换后文件: {result['data']['transformed_file']}")
        print(f"最终文件: {result['data']['final_file']}")
    else:
        print(f"处理失败: {result['message']}")

if __name__ == "__main__":
    local_handler()