"""hd 品牌处理器 — 兼容入口（委托至 src.processors.hd.HdProcessor）。"""
from src.processors.hd import HdProcessor

if __name__ == "__main__":
    HdProcessor().run()
