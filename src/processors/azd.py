"""azd 品牌处理器。"""
from src.processors.base_processor import BaseProcessor


class AzdProcessor(BaseProcessor):
    SOURCE_FILE = "610azd.xlsx"
    TEMPLATE_FILE = "./temp/prov_azdTemp.xlsm"
    CHUNK_SIZE = 500
