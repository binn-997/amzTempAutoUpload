"""std 品牌处理器。"""
from typing import List, Optional

from src.processors.base_processor import BaseProcessor


class StdProcessor(BaseProcessor):
    SOURCE_FILE = "521std.xlsx"
    TEMPLATE_FILE = "./temp/prov_stdTemp.xlsm"
    CHUNK_SIZE = 500
    SIZE_CLASS_COL: Optional[str] = None
    SIZE_TARGET_COLS: List[str] = ["CH", "FC"]
    PARENT_CLEAR_LETTERS: List[str] = [
        "V", "W", "AA", "AB",
        "BY", "BZ", "CF", "CG", "CH", "FC",
    ]
