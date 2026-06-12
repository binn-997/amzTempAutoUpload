"""std 品牌处理器 — 兼容入口（委托至 src.processors.std.StdProcessor）。"""
from src.processors.std import StdProcessor

if __name__ == "__main__":
    StdProcessor().run()
