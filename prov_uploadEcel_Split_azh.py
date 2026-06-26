"""azh（Anzug Herren 男士西装）品牌处理器 — 兼容入口。

用法：
    python prov_uploadEcel_Split_azh.py

此文件为向后兼容入口。内部从 config/categories.yaml 读取 azh 的默认配置，
委托通用 BaseProcessor 执行完整流水线。

如需指定非默认源文件，请使用 CLI 入口：
    python src/main.py process --type azh --source 其他文件.xlsx
"""

import os
import sys

_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from src.config_loader import (
    ConfigError,
    get_default_config_path,
    load_config,
    resolve_category_config,
)
from src.processors.base_processor import (
    BaseProcessor,
    IOFailure,
    ProcessingError,
)

_TYPE_KEY = "azh"

if __name__ == "__main__":
    try:
        config_path = get_default_config_path()
        config = load_config(config_path)
        cat_config = resolve_category_config(config, _TYPE_KEY)
        processor = BaseProcessor(cat_config)
        processor.run()
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
