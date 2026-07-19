"""
配置文件加载工具。

支持 YAML（.yaml/.yml）和 JSON（.json）格式。
YAML 需要 PyYAML 库（pip install pyyaml）；JSON 使用标准库。
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from copy import deepcopy
from typing import Any, Dict, List

# 修复 Windows 终端 GBK 编码问题
if sys.stdout.encoding != "utf-8":
    try:
        reconfigure = getattr(sys.stdout, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


class ConfigError(Exception):
    """配置文件加载或解析错误。"""


@dataclass(frozen=True)
class ConfigDocument:
    """经过基础结构校验的配置文档。"""

    defaults: dict[str, Any]
    source_columns: dict[str, Any]
    template_columns: dict[str, Any]
    categories: dict[str, dict[str, Any]]

    def as_dict(self) -> dict[str, Any]:
        return {
            "defaults": self.defaults,
            "source_columns": self.source_columns,
            "template_columns": self.template_columns,
            "categories": self.categories,
        }


def validate_config(config: Any) -> ConfigDocument:
    """校验配置的公共结构，尽早报告可操作的错误。"""
    if not isinstance(config, dict):
        raise ConfigError("配置根节点必须是对象")
    required = ("defaults", "source_columns", "template_columns", "categories")
    missing = [key for key in required if key not in config]
    if missing:
        raise ConfigError(f"配置缺少必填项: {', '.join(missing)}")
    categories = config["categories"]
    if not isinstance(categories, dict) or not categories:
        raise ConfigError("categories 必须是非空对象")
    for key, value in categories.items():
        if not isinstance(key, str) or not key.strip() or not isinstance(value, dict):
            raise ConfigError("每个产品类型必须是非空字符串键及对象值")
    for column in config["template_columns"].get("simple", {}).values():
        if not isinstance(column, str) or not column.isalpha():
            raise ConfigError(f"无效模板列字母: {column!r}")
    return ConfigDocument(
        defaults=dict(config["defaults"]),
        source_columns=dict(config["source_columns"]),
        template_columns=dict(config["template_columns"]),
        categories={key: dict(value) for key, value in categories.items()},
    )


def _deep_merge(base: dict, override: dict) -> dict:
    """深度合并两个字典。

    规则：
    - 标量值：override 覆盖 base（除非 override 为 None）
    - 字典值：递归合并
    - 列表值：override 完全替换 base
    """
    result = deepcopy(base)

    for key, value in override.items():
        if key not in result:
            result[key] = deepcopy(value)
        elif isinstance(value, dict) and isinstance(result[key], dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)

    return result


def load_config(config_path: str) -> dict:
    """加载配置文件，根据扩展名自动选择解析器。

    Args:
        config_path: 配置文件路径（.yaml / .yml / .json）

    Returns:
        解析后的配置字典。

    Raises:
        FileNotFoundError: 配置文件不存在。
        ConfigError: 配置格式错误或 YAML 库未安装。
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    ext = os.path.splitext(config_path)[1].lower()

    if ext in (".yaml", ".yml"):
        try:
            import yaml
        except ImportError:
            raise ConfigError(
                "读取 YAML 配置文件需要 PyYAML 库。\n"
                "请运行: pip install pyyaml\n"
                "或改用 JSON 格式配置文件（.json 扩展名）"
            )
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return validate_config(yaml.safe_load(f)).as_dict()
        except yaml.YAMLError as e:
            raise ConfigError(f"YAML 解析失败 [{config_path}]: {e}") from e

    elif ext == ".json":
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return validate_config(json.load(f)).as_dict()
        except json.JSONDecodeError as e:
            raise ConfigError(f"JSON 解析失败 [{config_path}]: {e}") from e

    else:
        raise ConfigError(
            f"不支持的配置文件格式: {ext}。"
            f"请使用 .yaml / .yml / .json 扩展名。"
        )


def resolve_category_config(
    config: dict,
    category_key: str,
) -> dict:
    """从配置中解析指定产品类型的完整配置（合并 defaults）。

    Args:
        config: load_config() 返回的完整配置字典。
        category_key: 产品类型标识（如 "azd"、"th"）。

    Returns:
        合并后的产品类型配置字典，可直接传入 BaseProcessor。

    Raises:
        ConfigError: 产品类型不存在。
    """
    categories: dict = config.get("categories", {})
    if category_key not in categories:
        available = ", ".join(sorted(categories.keys()))
        raise ConfigError(
            f"未知产品类型: {category_key}。"
            f"可用类型: {available}"
        )

    defaults = config.get("defaults", {})
    category_config = categories[category_key]

    # 深度合并：category_config 覆盖 defaults
    resolved = _deep_merge(defaults, category_config)

    # 注入共享的源列映射和模板列映射
    resolved["source_columns"] = deepcopy(config.get("source_columns", {}))
    resolved["template_columns"] = deepcopy(config.get("template_columns", {}))

    # 确保必要字段存在
    resolved.setdefault("title_prefix", "")
    resolved.setdefault("include_final_keywords", True)
    resolved.setdefault("output_dir", "./prov_output")
    resolved.setdefault("output_suffix", "_prov")
    resolved.setdefault("chunk_size", 500)

    return resolved


def list_categories(config: dict) -> List[Dict[str, str]]:
    """列出配置文件中所有可用的产品类型。

    Args:
        config: load_config() 返回的完整配置字典。

    Returns:
        [{"key": "azd", "description": "Anzug Damen（女士西装）"}, ...]
    """
    categories = config.get("categories", {})
    result = []
    for key, cat in categories.items():
        result.append({
            "key": key,
            "description": cat.get("description", key),
        })
    result.sort(key=lambda x: x["key"])
    return result


def find_project_root(start_path: str | None = None) -> str:
    """查找项目根目录（包含 config/ 目录的父目录）。

    从 start_path 向上搜索，直到找到包含 config/ 目录的位置。
    如果 start_path 为 None，则使用本文件的所在目录的父目录。

    Args:
        start_path: 起始搜索路径，默认为本文件所在目录的父目录。

    Returns:
        项目根目录的绝对路径。

    Raises:
        FileNotFoundError: 未找到项目根目录。
    """
    if start_path is None:
        start_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    current = os.path.abspath(start_path)
    for _ in range(10):  # 最多向上搜索 10 层
        config_dir = os.path.join(current, "config")
        if os.path.isdir(config_dir):
            return current
        parent = os.path.dirname(current)
        if parent == current:  # 已到文件系统根
            break
        current = parent

    raise FileNotFoundError(
        "无法找到项目根目录（未发现 config/ 目录）。"
        f"起始搜索路径: {start_path}"
    )


def get_default_config_path() -> str:
    """获取随安装包分发的默认 YAML 配置路径。"""
    return os.path.join(os.path.dirname(__file__), "config", "categories.yaml")
