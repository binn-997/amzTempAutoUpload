"""azd 品牌处理器 — 兼容入口（委托至 src.processors.azd.AzdProcessor）。"""
from src.processors.azd import AzdProcessor

if __name__ == "__main__":
    AzdProcessor().run()
