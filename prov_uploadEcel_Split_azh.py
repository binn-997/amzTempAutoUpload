"""azh 品牌处理器 — 兼容入口（委托至 src.processors.azh.AzhProcessor）。"""
from src.processors.azh import AzhProcessor

if __name__ == "__main__":
    AzhProcessor().run()
