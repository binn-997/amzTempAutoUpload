"""品牌处理器模块。每个文件只需定义配置，逻辑由 BaseProcessor 提供。"""

from src.processors.base_processor import BaseProcessor, ProcessingError

__all__ = ["BaseProcessor", "ProcessingError"]
