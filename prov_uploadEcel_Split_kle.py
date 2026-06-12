"""kle 品牌处理器 — 兼容入口（委托至 src.processors.kle.KleProcessor）。"""
from src.processors.kle import KleProcessor

if __name__ == "__main__":
    KleProcessor().run()
