"""自定义异常类模块"""

class DeliveryPlanError(Exception):
    """到货计划处理基础异常类"""
    pass

class FileOperationError(DeliveryPlanError):
    """文件操作相关异常"""
    pass

class ExcelOperationError(DeliveryPlanError):
    """Excel操作相关异常"""
    pass

class DataTransformError(DeliveryPlanError):
    """数据转换相关异常"""
    pass

class ConfigurationError(DeliveryPlanError):
    """配置相关异常"""
    pass
