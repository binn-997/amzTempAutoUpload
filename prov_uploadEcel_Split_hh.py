"""hh 品牌处理器 — 兼容入口（委托至 src.processors.hh.HhProcessor）。"""
from src.processors.hh import HhProcessor

if __name__ == "__main__":
    HhProcessor().run()
