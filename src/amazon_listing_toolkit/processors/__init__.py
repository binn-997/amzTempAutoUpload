"""品牌处理器模块。

产品类型已完全配置化（config/categories.yaml），不再需要子类文件。
BaseProcessor 通过配置字典动态创建实例。

产品差异仅由 YAML 配置表达。
"""

from .base_processor import BaseProcessor, ProcessingError, IOFailure

__all__ = ["BaseProcessor", "ProcessingError", "IOFailure"]
