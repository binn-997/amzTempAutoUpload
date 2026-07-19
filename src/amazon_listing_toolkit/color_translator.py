"""
颜色翻译与变体检测模块 — 英→德翻译 + 同款变体 V1/V2 标注。

支持自定义颜色映射表，可按 parent_sku 分组进行变体检测。
模块不触碰文件系统，Excel I/O 由 CLI 层负责。

用法:
    from amazon_listing_toolkit.color_translator import ColorTranslator, quick_translate

    ct = ColorTranslator()
    result = ct.process(df, color_col=6, size_col=7, parent_sku_col=2)
    for idx, new_color in result.index_to_new_color.items():
        df.iloc[idx, 6] = new_color

    # 便捷函数 — 仅翻译，不检测变体
    df = quick_translate(df)
"""

from __future__ import annotations

import logging
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


# ==================== 常量 ====================

COLOR_TRANSLATIONS: Dict[str, str] = {
    "Red": "Rot",
    "Beige": "Beige",
    "Black": "Schwarz",
    "Coffee": "Kaffee",
    "Navy": "Marineblau",
    "Orange": "Orange",
    "White": "Weiß",
    "Blue": "Blau",
    "Light Yellow": "Hellgelb",
    "Dark Green": "Dunkelgrün",
    "Pink": "Rosa",
    "Light Pink": "Hellrosa",
    "Light Blue": "Hellblau",
    "Hot Pink": "Knallrosa",
    "Mint Green": "Minzgrün",
    "Grey": "Grau",
    "Dark Gray": "Dunkelgrau",
    "Light Gray": "Hellgrau",
    "Yellow": "Gelb",
    "Army Green": "Armeegrün",
    "Purple": "Lila",
    "Green": "Grün",
    "Brick Red": "Ziegelrot",
    "Dark Blue": "Dunkelblau",
    "Wine": "Weinrot",
    "Multicolour": "Mehrfarbig",
    "Light Green": "Hellgrün",
    "Light Purple": "Helllila",
    "Rose Gold": "Roségold",
    "Sky Blue": "Himmelblau",
    "Brown": "Braun",
    "Khaki": "Khaki",
    "Camouflage": "Tarnfarben",
    "BU": "BU",
    "GN": "GN",
    "Watermelon Red": "Wassermelonenrot",
    "Dark Purple": "Dunkellila",
    "Gold": "Gold",
}


# ==================== 数据类 ====================


@dataclass
class ColorGroup:
    """同一 parent_sku 分组内连续相同颜色的行集合。

    Args:
        color: 颜色名称（未经翻译的原始值）。
        sizes: 该颜色组内各行的尺码列表（保持原始顺序）。
        row_indices: 该颜色组内各行在 DataFrame 中的行索引。
    """

    color: str
    sizes: List[str] = field(default_factory=list)
    row_indices: List[int] = field(default_factory=list)


@dataclass
class TranslateResult:
    """颜色翻译 + 变体检测的完整结果。

    Args:
        index_to_new_color: 行索引 → 新颜色值（已翻译 + 已加 V 后缀）的映射。
        total_rows: 参与处理的总行数。
        total_translated: 实际执行了翻译的颜色组数（英→德有变化）。
        total_variant_labeled: 被添加 V1/V2 后缀的行数。
    """

    index_to_new_color: Dict[int, str] = field(default_factory=dict)
    total_rows: int = 0
    total_translated: int = 0
    total_variant_labeled: int = 0


# ==================== 核心翻译器 ====================


class ColorTranslator:
    """颜色翻译器 — 英→德翻译 + 同款变体消歧。

    特性:
        - 支持自定义颜色映射表（默认使用内置 45 条英→德映射）
        - 按 parent_sku 分组进行变体检测（替代旧版的空行区块分割）
        - 变体签名 = (翻译后颜色, 尺码元组)，重复签名依次加 V1/V2 后缀
        - 非字符串颜色值（NaN / 数字）原样保留

    用法:
        ct = ColorTranslator()

        # 基础翻译
        ct.translate("Blue")   # → "Blau"
        ct.translate("Black1") # → "Schwarz1"

        # 处理 DataFrame
        result = ct.process(df, color_col=6, size_col=7, parent_sku_col=2)
    """

    def __init__(self, translations: Optional[Dict[str, str]] = None):
        """初始化翻译器。

        Args:
            translations: 自定义颜色映射表（英→德）。
                          若为 None，使用内置 COLOR_TRANSLATIONS。
        """
        self._translations = translations if translations is not None else dict(COLOR_TRANSLATIONS)

    # ── 属性 ──

    @property
    def translations(self) -> Dict[str, str]:
        """返回当前生效的颜色映射表（拷贝）。"""
        return dict(self._translations)

    # ── 翻译方法 ──

    def translate(self, color_name: str) -> str:
        """将英文颜色名翻译为德文。

        翻译流程:
            1. 精确匹配 COLOR_TRANSLATIONS
            2. 剥离尾部数字后缀后翻译基础部分，再拼回数字后缀
               （例如 "Black1" → base="Black" → "Schwarz" → "Schwarz1"）
            3. 无法匹配则原样返回

        Args:
            color_name: 英文颜色名（可带数字后缀，如 "Blue2"）。

        Returns:
            德文颜色名。若无法翻译则返回原值。
        """
        # 非字符串安全处理
        if not isinstance(color_name, str):
            return color_name

        color_name = color_name.strip()

        # 1. 精确匹配
        if color_name in self._translations:
            return self._translations[color_name]

        # 2. 剥离尾部数字后缀再匹配
        m = re.match(r"^(.+?)(\d+)$", color_name)
        if m:
            base = m.group(1)
            suffix = m.group(2)
            if base in self._translations:
                return self._translations[base] + suffix

        # 3. 都不匹配，原样返回
        return color_name

    # ── 内部方法 ──

    @staticmethod
    def _extract_color_groups(
        df: pd.DataFrame,
        color_col: int,
        size_col: int,
    ) -> List[ColorGroup]:
        """从 DataFrame 子集中提取连续同色的颜色组。

        遍历给定行，将连续相同颜色值的行聚合为一个 ColorGroup。
        遇到颜色变化时开始新组。

        Args:
            df: 已筛选好的 DataFrame 子集（某 parent_sku 下的所有子体行）。
            color_col: 颜色列的 0-based 索引。
            size_col: 尺码列的 0-based 索引。

        Returns:
            ColorGroup 列表，保持原始行顺序。
        """
        groups: List[ColorGroup] = []
        current_color = None
        current_sizes: List[str] = []
        current_indices: List[int] = []

        for idx, row in df.iterrows():
            color_val = row.iloc[color_col]
            size_val = row.iloc[size_col]

            # 跳过颜色为空的行（父体行）
            if pd.isna(color_val) or str(color_val).strip() == "":
                continue

            color_str = str(color_val).strip()
            size_str = str(size_val).strip() if pd.notna(size_val) else ""

            if color_str != current_color:
                if current_color is not None:
                    groups.append(ColorGroup(
                        color=current_color,
                        sizes=current_sizes[:],
                        row_indices=current_indices[:],
                    ))
                current_color = color_str
                current_sizes = [size_str]
                current_indices = [idx]
            else:
                current_sizes.append(size_str)
                current_indices.append(idx)

        # 收尾最后一个组
        if current_color is not None:
            groups.append(ColorGroup(
                color=current_color,
                sizes=current_sizes[:],
                row_indices=current_indices[:],
            ))

        return groups

    @staticmethod
    def _detect_and_label_variants(
        groups: List[ColorGroup],
    ) -> tuple[Dict[int, str], int]:
        """对颜色组进行变体检测，为重复签名添加 V1/V2 后缀。

        签名 = (颜色名, tuple(尺码列表))。
        - 签名出现 1 次 → 不改名
        - 签名出现 N 次 → 依次标注 "ColorV1", "ColorV2", ..., "ColorVN"

        注意：传入的 groups 应当已经完成颜色翻译。

        Args:
            groups: 已完成翻译的 ColorGroup 列表。

        Returns:
            (行索引 → 带变体后缀的颜色名 的映射, 被添加 V 后缀的行数)。
        """
        # 按签名分组
        signature_map: Dict[tuple, List[List[int]]] = defaultdict(list)
        for g in groups:
            sig = (g.color, tuple(g.sizes))
            signature_map[sig].append(g.row_indices)

        index_to_color: Dict[int, str] = {}
        variant_labeled = 0

        for (color, _), index_lists in signature_map.items():
            n = len(index_lists)
            if n == 1:
                for idx in index_lists[0]:
                    index_to_color[idx] = color
            else:
                for i, indices in enumerate(index_lists, 1):
                    label = f"{color}V{i}"
                    variant_labeled += len(indices)
                    for idx in indices:
                        index_to_color[idx] = label

        return index_to_color, variant_labeled

    # ── 主处理方法 ──

    def process(
        self,
        df: pd.DataFrame,
        color_col: int = 6,
        size_col: int = 7,
        parent_sku_col: int = 2,
        enable_translation: bool = True,
        enable_variant: bool = True,
    ) -> TranslateResult:
        """对 DataFrame 执行颜色翻译和变体检测。

        处理流程:
            1. 按 parent_sku_col 列将数据分组（每个 parent_sku 值 = 一个产品变体块）
            2. 每个块内：跳过父体行（颜色为空），在子体行上提取颜色组
            3. 翻译各颜色组的颜色名（可选）
            4. 检测块内重复签名，添加 V1/V2 后缀（可选）
            5. 汇总所有块的结果

        Args:
            df: 源 DataFrame（header=None 读取的原始数据）。
            color_col: 颜色列的 0-based 索引（默认 6 = G 列）。
            size_col: 尺码列的 0-based 索引（默认 7 = H 列）。
            parent_sku_col: parent_sku 列的 0-based 索引（默认 2 = C 列）。
            enable_translation: 是否执行颜色翻译。
            enable_variant: 是否执行变体检测。

        Returns:
            TranslateResult 包含完整的行索引→新颜色映射和统计信息。
        """
        index_to_new_color: Dict[int, str] = {}
        total_rows = 0
        total_translated = 0
        total_variant_labeled = 0

        # 1. 获取所有唯一的 parent_sku 值
        #    NaN 值视为独立分组（每行自成一块）
        parent_col = df.iloc[:, parent_sku_col]
        unique_parents = parent_col.dropna().unique().tolist()

        # 记录已通过 parent_sku 处理的索引
        processed_indices: set = set()

        for parent_sku_val in unique_parents:
            # 获取该 parent_sku 下的所有行
            mask = parent_col == parent_sku_val
            block_df = df[mask]
            block_indices = set(block_df.index.tolist())
            processed_indices.update(block_indices)

            # 提取颜色组（自动跳过颜色为空的行）
            groups = self._extract_color_groups(block_df, color_col, size_col)
            if not groups:
                continue

            total_rows += sum(len(g.row_indices) for g in groups)

            # 翻译颜色
            if enable_translation:
                for g in groups:
                    translated = self.translate(g.color)
                    if translated != g.color:
                        total_translated += 1
                    g.color = translated

            # 变体检测
            if enable_variant:
                local_map, variant_count = self._detect_and_label_variants(groups)
                total_variant_labeled += variant_count
                for idx, color_val in local_map.items():
                    index_to_new_color[idx] = color_val
            else:
                for g in groups:
                    for idx in g.row_indices:
                        index_to_new_color[idx] = g.color

        # 2. 处理 parent_sku 为 NaN 的行（每行自成一块，不做变体检测）
        nan_mask = parent_col.isna()
        nan_indices = set(df[nan_mask].index.tolist()) - processed_indices

        for idx in nan_indices:
            color_val = df.iloc[idx, color_col]
            if pd.isna(color_val) or str(color_val).strip() == "":
                continue
            color_str = str(color_val).strip()

            if enable_translation:
                translated = self.translate(color_str)
                if translated != color_str:
                    total_translated += 1
                color_str = translated

            index_to_new_color[idx] = color_str
            total_rows += 1

        return TranslateResult(
            index_to_new_color=index_to_new_color,
            total_rows=total_rows,
            total_translated=total_translated,
            total_variant_labeled=total_variant_labeled,
        )


# ==================== 便捷函数 ====================

_DEFAULT_TRANSLATOR: Optional[ColorTranslator] = None


def quick_translate(
    df: pd.DataFrame,
    color_col: int = 6,
) -> pd.DataFrame:
    """快速翻译 DataFrame 中的颜色列（仅翻译，不检测变体）。

    返回新 DataFrame（拷贝），不修改原数据。

    Args:
        df: 源 DataFrame。
        color_col: 颜色列的 0-based 索引。

    Returns:
        颜色列已翻译的新 DataFrame。
    """
    global _DEFAULT_TRANSLATOR
    if _DEFAULT_TRANSLATOR is None:
        _DEFAULT_TRANSLATOR = ColorTranslator()
    ct = _DEFAULT_TRANSLATOR
    result_df = df.copy()
    for idx in range(len(result_df)):
        val = result_df.iloc[idx, color_col]
        if pd.notna(val) and isinstance(val, str):
            result_df.iloc[idx, color_col] = ct.translate(val)
    return result_df


# ==================== 自测入口 ====================

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    print("\n" + "=" * 50)
    print("  ColorTranslator Self-Test")
    print("=" * 50)

    ct = ColorTranslator()

    # ── 测试 translate() ──
    print("\n--- translate() 单元测试 ---")
    tests = [
        ("Blue", "Blau"),
        ("Dark Green", "Dunkelgrün"),
        ("Black1", "Schwarz1"),
        ("Light Blue3", "Hellblau3"),
        ("Schwarz", "Schwarz"),          # 已是德语，幂等
        ("UnknownColor", "UnknownColor"), # 未知，原样返回
        ("Red", "Rot"),
        ("Pink", "Rosa"),
        ("Watermelon Red", "Wassermelonenrot"),
    ]
    for input_val, expected in tests:
        translated = ct.translate(input_val)
        status = "PASS" if translated == expected else f"FAIL (got {translated})"
        print(f"  translate({input_val!r:25s}) = {translated!r:25s}  [{status}]")
        assert translated == expected, f"translate({input_val!r}) failed: {translated} != {expected}"

    # ── 测试非字符串处理 ──
    print("\n--- translate() 非字符串测试 ---")
    nan_val = float("nan")
    result_nan = ct.translate(nan_val)  # type: ignore[arg-type]
    assert pd.isna(result_nan), f"NaN should remain NaN, got {result_nan}"
    assert ct.translate(123) == 123  # type: ignore[arg-type]
    print("  NaN / 数字 → 原样返回 [PASS]")

    # ── 测试 _extract_color_groups ──
    print("\n--- _extract_color_groups() 测试 ---")
    # 构造足够宽的数据以访问列索引 6/7
    import numpy as np
    test_data = pd.DataFrame(index=range(7), columns=range(8))
    test_data.iloc[:, 6] = ["Pink", "Pink", "Pink", "Black", "Black", np.nan, "Blue"]
    test_data.iloc[:, 7] = ["S", "M", "L", "XL", "XXL", np.nan, "M"]
    groups = ct._extract_color_groups(test_data, color_col=6, size_col=7)
    assert len(groups) == 3, f"Expected 3 groups, got {len(groups)}"
    assert groups[0].color == "Pink" and groups[0].sizes == ["S", "M", "L"]
    assert groups[1].color == "Black" and groups[1].sizes == ["XL", "XXL"]
    # 第 6 行颜色为 None，应被跳过
    assert groups[2].color == "Blue" and groups[2].sizes == ["M"]
    print("  3 组提取正确: Pink(S/M/L), Black(XL/XXL), Blue(M) [PASS]")

    # ── 测试变体检测 ──
    print("\n--- _detect_and_label_variants() 测试 ---")
    variant_groups = [
        ColorGroup(color="Schwarz", sizes=["L", "M", "S"], row_indices=[1, 2, 3]),
        ColorGroup(color="Schwarz", sizes=["L", "M", "S"], row_indices=[4, 5, 6]),
        ColorGroup(color="Rot", sizes=["XL"], row_indices=[7]),
    ]
    label_map, variant_count = ct._detect_and_label_variants(variant_groups)
    assert label_map[1] == "SchwarzV1"
    assert label_map[4] == "SchwarzV2"
    assert label_map[7] == "Rot"  # 唯一签名，不加后缀
    assert variant_count == 6  # SchwarzV1(3行) + SchwarzV2(3行)
    print("  Schwarz 重复 → V1/V2, Rot 唯一 → 不变 [PASS]")

    # ── 测试 process() 完整流程 ──
    print("\n--- process() 完整流程测试 ---")
    df_test = pd.DataFrame(index=range(5), columns=range(8))
    df_test.iloc[:, 1] = ["SKU1", "SKU2", "SKU3", "SKU4", "SKU5"]     # sku
    df_test.iloc[:, 2] = ["P1", "P1", "P1", "P2", "P2"]                # parent_sku
    df_test.iloc[:, 6] = ["Blue", "Blue", "Red", "Blue", "Blue"]       # color (G)
    df_test.iloc[:, 7] = ["S", "M", "L", "S", "M"]                     # size (H)
    result = ct.process(df_test, color_col=6, size_col=7, parent_sku_col=2)
    # P1: Blue(S,M) → Blau(S,M), Red(L) → Rot(L)
    # P2: Blue(S,M) → Blau(S,M) — 与 P1 不同的 parent，不触发变体
    assert result.total_rows == 5
    assert result.total_translated >= 2  # Blue→Blau 至少 2组
    assert result.index_to_new_color[0] in ("Blau", "BlauV1")
    assert result.index_to_new_color[2] == "Rot"
    print(f"  total_rows={result.total_rows}, translated={result.total_translated}, "
          f"variant_labeled={result.total_variant_labeled} [PASS]")

    # ── 测试 disable flags ──
    print("\n--- process() 禁用标志测试 ---")
    result_no_trans = ct.process(df_test, color_col=6, size_col=7, parent_sku_col=2,
                                 enable_translation=False, enable_variant=True)
    assert result_no_trans.total_translated == 0
    print("  enable_translation=False → 0 翻译 [PASS]")

    result_no_var = ct.process(df_test, color_col=6, size_col=7, parent_sku_col=2,
                               enable_translation=True, enable_variant=False)
    assert result_no_var.total_variant_labeled == 0
    print("  enable_variant=False → 0 变体标注 [PASS]")

    # ── 测试 quick_translate ──
    print("\n--- quick_translate() 测试 ---")
    df_quick = pd.DataFrame(index=range(3), columns=range(7))
    df_quick.iloc[:, 6] = ["Blue", "Red", "Black"]
    df_result = quick_translate(df_quick, color_col=6)
    assert df_result.iloc[0, 6] == "Blau"
    assert df_result.iloc[1, 6] == "Rot"
    assert df_result.iloc[2, 6] == "Schwarz"
    assert df_quick.iloc[0, 6] == "Blue"  # 原 df 未被修改
    print("  翻译正确 & 原 DataFrame 未被修改 [PASS]")

    print("\n" + "=" * 50)
    print("  All tests passed.")
    print("=" * 50 + "\n")
