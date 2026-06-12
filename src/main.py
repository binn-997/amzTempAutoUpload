"""
统一 CLI 入口 — 品牌处理器。

用法:
    python src/main.py --brand azd
    python src/main.py --brand hd --chunk 800 --dry-run
    python src/main.py --brand azh --input 610azh.xlsx --output-dir ./out
"""

from __future__ import annotations

import argparse
import sys
from typing import Type

from src.processors.azd import AzdProcessor
from src.processors.azh import AzhProcessor
from src.processors.base_processor import BaseProcessor, ProcessingError
from src.processors.hd import HdProcessor
from src.processors.hh import HhProcessor
from src.processors.kle import KleProcessor
from src.processors.std import StdProcessor
from src.processors.td import TdProcessor
from src.processors.th import ThProcessor

# 品牌名 → 处理器类 注册表
BRAND_REGISTRY: dict[str, Type[BaseProcessor]] = {
    "azd": AzdProcessor,
    "azh": AzhProcessor,
    "hd": HdProcessor,
    "hh": HhProcessor,
    "kle": KleProcessor,
    "std": StdProcessor,
    "td": TdProcessor,
    "th": ThProcessor,
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Amazon 商品数据 → 模板填充处理器",
    )
    parser.add_argument(
        "--brand",
        required=True,
        choices=list(BRAND_REGISTRY),
        help="品牌代码",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="源 Excel 文件（覆盖默认配置）",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="模板文件路径（覆盖默认配置）",
    )
    parser.add_argument(
        "--chunk",
        type=int,
        default=None,
        help="分块行数（覆盖默认配置）",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="输出目录（覆盖默认配置）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只做预处理分析，不写文件",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    processor_cls = BRAND_REGISTRY[args.brand]
    processor = processor_cls()

    # CLI 参数覆盖默认配置
    if args.input:
        processor.SOURCE_FILE = args.input
    if args.template:
        processor.TEMPLATE_FILE = args.template
    if args.chunk:
        processor.CHUNK_SIZE = args.chunk
    if args.output_dir:
        processor.OUTPUT_DIR = args.output_dir

    try:
        if args.dry_run:
            processor.dry_run()
        else:
            processor.run()
    except ProcessingError as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"❌ 未预期的错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
