"""th 品牌处理器 — 兼容入口（委托至 src.processors.th.ThProcessor）。"""
from src.processors.th import ThProcessor

if __name__ == "__main__":
    ThProcessor().run()
