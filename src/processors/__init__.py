"""品牌处理器模块。

产品类型已完全配置化（config/categories.yaml），不再需要子类文件。
BaseProcessor 通过配置字典动态创建实例。

旧版兼容入口（prov_uploadEcel_Split_*.py）仍可用，
它们从配置文件读取默认参数后委托 BaseProcessor 执行。
"""

from src.processors.base_processor import BaseProcessor, ProcessingError, IOFailure

__all__ = ["BaseProcessor", "ProcessingError", "IOFailure"]
