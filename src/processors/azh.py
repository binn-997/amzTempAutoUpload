"""azh 品牌处理器。"""
from src.processors.base_processor import BaseProcessor


class AzhProcessor(BaseProcessor):
    SOURCE_FILE = "507azh.xlsx"
    TEMPLATE_FILE = "./temp/prov_azhTemp.xlsm"
    CHUNK_SIZE = 500
