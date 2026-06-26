"""
统一 CLI 入口 — 配置驱动的品牌处理器。

用法:
    # 推荐：作为模块运行
    python -m src.main process --type azd

    # 也可直接运行脚本
    python src/main.py process --type azd

    # 覆盖源文件和分块大小
    python src/main.py process --type azd --source 624azh.xlsx --chunk 800

    # 使用自定义配置文件
    python src/main.py process --type hd --config my_config.yaml

    # 仅预览，不写入文件
    python src/main.py process --type th --dry-run

    # 列出所有可用产品类型
    python src/main.py list-types
    python src/main.py list-types --config my_config.yaml
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import Optional

# 修复 Windows 终端 GBK 编码问题（支持 emoji 和中文输出）
if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 确保项目根目录在 sys.path 中（支持直接 python src/main.py 执行）
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from src.config_loader import (
    ConfigError,
    get_default_config_path,
    list_categories,
    load_config,
    resolve_category_config,
)
from src.processors.base_processor import (
    BaseProcessor,
    IOFailure,
    ProcessingError,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Amazon 商品数据 → 模板填充处理器（配置驱动）",
    )
    sub = parser.add_subparsers(dest="command", help="可用子命令")

    # ---- process 子命令 ----
    proc = sub.add_parser("process", help="执行数据处理流水线")
    proc.add_argument(
        "--type",
        required=True,
        help="产品类型标识（如 azd, hd, th）。必选。",
    )
    proc.add_argument(
        "--source",
        default=None,
        help="源 Excel 文件路径（覆盖配置文件中的默认值）",
    )
    proc.add_argument(
        "--template",
        default=None,
        help="模板文件路径（覆盖配置文件中的默认值）",
    )
    proc.add_argument(
        "--chunk",
        type=int,
        default=None,
        help="分块行数（覆盖配置文件中的默认值）",
    )
    proc.add_argument(
        "--output-dir",
        default=None,
        help="输出目录（覆盖配置文件中的默认值）",
    )
    proc.add_argument(
        "--config",
        default=None,
        help="配置文件路径（默认: config/categories.yaml）",
    )
    proc.add_argument(
        "--dry-run",
        action="store_true",
        help="只做预处理分析，不写文件",
    )

    # ---- list-types 子命令 ----
    lst = sub.add_parser("list-types", help="列出所有可用产品类型")
    lst.add_argument(
        "--config",
        default=None,
        help="配置文件路径（默认: config/categories.yaml）",
    )

    return parser


def cmd_process(args: argparse.Namespace) -> None:
    """执行 process 子命令。"""
    # 1. 加载配置文件
    config_path = args.config or get_default_config_path()
    print(f"📋 配置文件: {config_path}")
    config = load_config(config_path)

    # 2. 解析产品类型配置
    cat_config = resolve_category_config(config, args.type)
    description = cat_config.get("description", args.type)
    print(f"🏷️  产品类型: {args.type} ({description})")

    # 3. CLI 参数覆盖配置文件默认值
    if args.source:
        cat_config["default_source"] = args.source
    if args.template:
        cat_config["template_file"] = args.template
    if args.chunk:
        cat_config["chunk_size"] = args.chunk
    if args.output_dir:
        cat_config["output_dir"] = args.output_dir

    # 4. 创建处理器并执行
    processor = BaseProcessor(cat_config)

    if args.dry_run:
        processor.dry_run()
    else:
        processor.run()


def cmd_list_types(args: argparse.Namespace) -> None:
    """执行 list-types 子命令。"""
    config_path = args.config or get_default_config_path()
    print(f"📋 配置文件: {config_path}")
    config = load_config(config_path)

    categories = list_categories(config)
    if not categories:
        print("（无可用产品类型）")
        return

    print(f"\n{'标识':<8} 描述")
    print("-" * 50)
    for cat in categories:
        print(f"{cat['key']:<8} {cat['description']}")


def main(argv: Optional[list[str]] = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        sys.exit(0)

    try:
        if args.command == "process":
            cmd_process(args)
        elif args.command == "list-types":
            cmd_list_types(args)
    except ConfigError as e:
        print(f"❌ 配置错误: {e}", file=sys.stderr)
        sys.exit(1)
    except ProcessingError as e:
        print(f"❌ 处理错误: {e}", file=sys.stderr)
        sys.exit(1)
    except IOFailure as e:
        print(f"❌ 文件错误: {e}", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"❌ 未预期的错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
