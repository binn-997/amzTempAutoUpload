"""td 品牌处理器 — 兼容入口（委托至 src.processors.td.TdProcessor）。"""
from src.processors.td import TdProcessor

if __name__ == "__main__":
    TdProcessor().run()
