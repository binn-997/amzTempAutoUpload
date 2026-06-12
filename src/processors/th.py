"""th 品牌处理器。"""
from typing import Dict, List

from src.processors.base_processor import BaseProcessor


class ThProcessor(BaseProcessor):
    SOURCE_FILE = "609WM2.xlsx"
    TEMPLATE_FILE = "./temp/prov_thTemp.xlsm"
    CHUNK_SIZE = 500
    PARENT_CLEAR_LETTERS: List[str] = [
        "V", "W", "X", "AA", "AB", "AC", "AD", "AE", "AG", "AH", "AZ",
        "BY", "BZ", "CF", "CG", "CH", "FC",
    ]
    SIZE_REPLACEMENTS: Dict[str, str] = {
        r"\bXXXL\b": "3XL",
        r"\bXXXXL\b": "4XL",
        r"\bXXXXXL\b": "5XL",
        r"\bXXXXXXL\b": "6XL",
        r"\bXXXXXXXL\b": "7XL",
        r"\bXXXXXXXXL\b": "8XL",
        r"\bXXXXXXXXXL\b": "9XL",
        r"\bL2\b": "XXL",
        r"\bL3\b": "3XL",
        r"\bL4\b": "4XL",
        r"\bL5\b": "5XL",
        r"\bone size\b": "Einheitsgröße",
    }
