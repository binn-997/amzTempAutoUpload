"""hd 品牌处理器。"""
from typing import List

from src.processors.base_processor import BaseProcessor


class HdProcessor(BaseProcessor):
    SOURCE_FILE = "610hd.xlsx"
    TEMPLATE_FILE = "./temp/prov_hdTemp.xlsx"
    CHUNK_SIZE = 1000
    SIZE_CLASS_COL = "BA"
    SIZE_TARGET_COLS: List[str] = ["BB", "CH", "FC"]
    PARENT_CLEAR_LETTERS: List[str] = [
        "V", "W", "X", "AA", "AB", "BA", "BB", "BC", "BF", "BG", "AZ",
        "BY", "BZ", "CF", "CG", "CH", "FC",
    ]
