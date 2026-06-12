"""td 品牌处理器。"""
from typing import List

from src.processors.base_processor import BaseProcessor


class TdProcessor(BaseProcessor):
    SOURCE_FILE = "611WMtd.xlsx"
    TEMPLATE_FILE = "./temp/prov_tdTemp.xlsm"
    CHUNK_SIZE = 1000
    PARENT_CLEAR_LETTERS: List[str] = [
        "V", "W", "X", "AA", "AB", "AC", "AD", "AE", "AG", "AH", "AZ",
        "BY", "BZ", "CF", "CG", "CH", "FC",
    ]
