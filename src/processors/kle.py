"""kle 品牌处理器。"""
from src.processors.base_processor import BaseProcessor


class KleProcessor(BaseProcessor):
    SOURCE_FILE = "609kle.xlsx"
    TEMPLATE_FILE = "./temp/prov_kleTemp.xlsm"
    CHUNK_SIZE = 1000
