import pandas as pd
import re
from collections import defaultdict
from datetime import datetime

# ===================== 配置模块 =====================
COLOR_TRANSLATIONS = {
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

INPUT_FILE = "data/1.xlsx"
OUTPUT_FILE = f"data/processed_1_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
SHEET_NAME = "Sheet1"


# ===================== 工具模块 =====================

def split_into_blocks(df):
    """
    按空行分割为独立区块，空行保留为标记区块。
    """
    blocks = []
    current_block = []

    for idx, row in df.iterrows():
        if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]):
            if current_block:
                blocks.append(current_block)
                current_block = []
            blocks.append([(idx, None)])
        else:
            current_block.append((idx, row))

    if current_block:
        blocks.append(current_block)

    return blocks


def extract_color_groups(block):
    """
    区块内按连续相同颜色聚合为「颜色组」。
    返回: [(color, [size, ...], [idx, ...]), ...]
    """
    groups = []
    current_color = None
    current_sizes = []
    current_indices = []

    for idx, row in block:
        if row is None:
            continue
        color = row.iloc[0]
        size  = row.iloc[1]

        if color != current_color:
            if current_color is not None:
                groups.append((current_color, current_sizes, current_indices))
            current_color = color
            current_sizes = [size]
            current_indices = [idx]
        else:
            current_sizes.append(size)
            current_indices.append(idx)

    if current_color is not None:
        groups.append((current_color, current_sizes, current_indices))

    return groups


def add_variant_labels(groups):
    """
    签名 = (color, tuple(sizes))
    - 唯一 → 不改名
    - 出现 N 次 → 依次 "ColorV1", "ColorV2", ...
    注意：传入的 color 应当已经翻译过，这里只负责查重加后缀。
    """
    signature_map = defaultdict(list)

    for color, sizes, indices in groups:
        sig = (color, tuple(sizes))
        signature_map[sig].append(indices)

    index_to_new_color = {}

    for (color, _), index_lists in signature_map.items():
        n = len(index_lists)
        if n == 1:
            for idx in index_lists[0]:
                index_to_new_color[idx] = color
        else:
            for i, indices in enumerate(index_lists, 1):
                label = f"{color}V{i}"
                for idx in indices:
                    index_to_new_color[idx] = label

    return index_to_new_color


def translate_color(color_name, translate=True):
    """
    翻译颜色名（支持带数字后缀的变体标签）。
    Blue       → Blau
    Dark Green → Dunkelgrün
    Black1     → Schwarz1   （先剥数字→翻译→拼回）
    """
    if not translate:
        return color_name

    # 安全处理非字符串值（如 NaN）
    if not isinstance(color_name, str):
        return color_name

    color_name = color_name.strip()

    # 1. 精确匹配
    if color_name in COLOR_TRANSLATIONS:
        return COLOR_TRANSLATIONS[color_name]

    # 2. 剥离尾部数字后缀再匹配（如 Black1 → Black，Red2 → Red）
    m = re.match(r"^(.+?)(\d+)$", color_name)
    if m:
        base = m.group(1)
        suffix = m.group(2)
        if base in COLOR_TRANSLATIONS:
            return COLOR_TRANSLATIONS[base] + suffix

    # 3. 都不匹配，原样返回
    return color_name


# ===================== 核心处理模块 =====================

def process_excel_file(input_file, output_file, sheet_name="Sheet1",
                       enable_translation=True):
    print(f"读取: {input_file}")

    xls = pd.ExcelFile(input_file)
    all_sheets = {}

    for sheet in xls.sheet_names:
        if sheet == sheet_name:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            original = df.copy()

            print(f"处理工作表: {sheet}  |  总行数: {len(df)}")

            blocks = split_into_blocks(df)
            print(f"数据区块: {len(blocks)}")

            global_map = {}
            for blk in blocks:
                if len(blk) == 1 and blk[0][1] is None:
                    continue
                groups = extract_color_groups(blk)
                # 1. 先翻译颜色
                translated_groups = [
                    (translate_color(color, enable_translation), sizes, indices)
                    for color, sizes, indices in groups
                ]
                # 2. 再查重加 V1/V2 后缀
                local_map = add_variant_labels(translated_groups)
                global_map.update(local_map)

            changed = 0
            for idx, new_color in global_map.items():
                if new_color != original.iloc[idx, 0]:
                    df.iloc[idx, 0] = new_color
                    changed += 1

            print(f"已修改: {changed} 行")
            all_sheets[sheet] = df
        else:
            all_sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)

    print(f"写入: {output_file}")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sn, df_out in all_sheets.items():
            df_out.to_excel(writer, sheet_name=sn, index=False, header=False)

    print("完成！")


# ===================== 主入口 =====================
if __name__ == "__main__":
    process_excel_file(
        input_file=INPUT_FILE,
        output_file=OUTPUT_FILE,
        sheet_name=SHEET_NAME,
        enable_translation=True,
    )