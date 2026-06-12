"""
品牌处理器基类 — 包含所有公共逻辑。

8 个品牌子类只需覆盖 ~5 个差异属性即可。
"""

from __future__ import annotations

import math
import os
import re
from typing import Any, Dict, List, Optional, Pattern

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet


class ProcessingError(Exception):
    """处理流程中的可恢复错误。"""


class BaseProcessor:
    """Amazon 商品数据→模板填充处理器的公共基类。

    子类只需覆盖以下属性：
        SOURCE_FILE, TEMPLATE_FILE, CHUNK_SIZE
        SIZE_CLASS_COL, INCLUDE_FINAL_KEYWORDS
        SIZE_TARGET_COLS, PARENT_CLEAR_LETTERS, SIZE_REPLACEMENTS
    """

    # ==================== 子类必须覆盖 ====================

    SOURCE_FILE: str = ""
    TEMPLATE_FILE: str = ""
    CHUNK_SIZE: int = 500

    # size_class 映射到哪一列（None = 不输出 size_class）
    SIZE_CLASS_COL: Optional[str] = "AD"

    # 是否输出 final_keywords 列
    INCLUDE_FINAL_KEYWORDS: bool = True

    # MULTI_MAP 中 size 对应的目标列
    SIZE_TARGET_COLS: List[str] = ["AE", "CH", "FC"]

    # 父体行需要清空的列字母列表
    PARENT_CLEAR_LETTERS: List[str] = [
        "V", "W", "X", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH",
        "BY", "BZ", "CF", "CG", "CH", "FC",
    ]

    # 尺码替换规则（key 为正则 pattern）
    SIZE_REPLACEMENTS: Dict[str, str] = {
        r"\bXXXL\b": "3XL",
        r"\bXXXXL\b": "4XL",
        r"\bXXXXXL\b": "5XL",
        r"\bXXXXXXL\b": "6XL",
        r"\bXXXXXXXL\b": "7XL",
        r"\bXXXXXXXXL\b": "8XL",
        r"\bXXXXXXXXXL\b": "9XL",
        r"\bone size\b": "Einheitsgröße",
    }

    # ==================== 公共常量（无需覆盖） ====================

    OUTPUT_SUFFIX: str = "_prov"
    MAX_SOURCE_COLS: int = 60
    START_ROW: int = 4
    MAX_IMAGES: int = 7
    TITLE_PREFIX: str = ""
    SHEET_NAME: str = "Vorlage"
    OUTPUT_DIR: str = "./prov_output"

    # --- 源数据列索引 (0-based) ---
    SRC: Dict[str, Any] = {
        "sku": 1,
        "parent_sku": 2,
        "title": 3,
        "color": 6,
        "size": 7,
        "package_weight": 9,
        "standprice": 38,
        "list_price_tax": 39,
        "generic_keywords": 45,
        "bullet1": 40,
        "bullet2": 41,
        "bullet3": 42,
        "bullet4": 43,
        "bullet5": 44,
        "images": [10, 11, 12, 13, 14, 15, 16],
    }

    # --- Bullet Points 映射 ---
    BULLET_SRC_PREFIX: str = "final_bullet"
    BULLET_TGT_COLS: List[str] = ["CO", "CP", "CQ", "CR", "CS"]

    # --- 图片映射 ---
    IMG_SRC_COLS: List[int] = [10, 11, 12, 13, 14, 15, 16]
    IMG_TGT_START: str = "BH"

    # --- 父子标识列 ---
    PC_COL: str = "BX"

    # ==================== 计算属性 ====================

    @property
    def parent_clear_idx(self) -> List[int]:
        """父体清空列的 1-based 索引列表（延迟计算）。"""
        return [column_index_from_string(c) for c in self.PARENT_CLEAR_LETTERS]

    @property
    def simple_map(self) -> Dict[str, str]:
        """一对一列映射：{DataFrame 列名: 目标 Excel 列字母}。"""
        m: Dict[str, str] = {
            "sku": "B",
            "parent_sku": "BY",
            "title": "G",
            "standprice": "V",
            "package_weight": "GD",
            "list_price_tax": "KP",
            "final_keywords": "CE",
        }
        if self.SIZE_CLASS_COL:
            m["size_class"] = self.SIZE_CLASS_COL
        if not self.INCLUDE_FINAL_KEYWORDS:
            del m["final_keywords"]
        return m

    @property
    def multi_map(self) -> Dict[str, List[str]]:
        """一对多列映射：{DataFrame 列名: [目标 Excel 列字母列表]}。"""
        return {
            "color": ["CF", "CG"],
            "size": self.SIZE_TARGET_COLS,
        }

    @property
    def _simple_idx(self) -> Dict[str, int]:
        return {k: column_index_from_string(v) for k, v in self.simple_map.items()}

    @property
    def _multi_idx(self) -> Dict[str, List[int]]:
        return {
            k: [column_index_from_string(c) for c in v]
            for k, v in self.multi_map.items()
        }

    @property
    def _bullet_idx(self) -> List[int]:
        return [column_index_from_string(c) for c in self.BULLET_TGT_COLS]

    @property
    def _img_start(self) -> int:
        return column_index_from_string(self.IMG_TGT_START)

    @property
    def _pc_idx(self) -> int:
        return column_index_from_string(self.PC_COL)

    # ==================== 数据预处理 ====================

    def preprocess_source_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """在内存中完成所有数据清洗与父子体继承。"""
        print("预处理源数据：清洗空值、替换尺码、继承父体数据...")

        # 列补齐 + 截断 + 空值填充
        if df.shape[1] < self.MAX_SOURCE_COLS:
            for i in range(df.shape[1], self.MAX_SOURCE_COLS):
                df[i] = ""
        df = df.iloc[:, : self.MAX_SOURCE_COLS].fillna("")
        df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)

        # 列重命名
        rename_map = {v: k for k, v in self.SRC.items() if isinstance(v, int)}
        df = df.rename(columns=rename_map)

        # title 加前缀
        df["title"] = self.TITLE_PREFIX + df["title"].astype(str)

        # 尺码替换
        for pattern, replacement in self.SIZE_REPLACEMENTS.items():
            df["size"] = (
                df["size"]
                .astype(str)
                .str.replace(pattern, replacement, regex=True, flags=re.IGNORECASE)
            )

        # size_class
        df["size_class"] = df["size"].apply(
            lambda x: "Numerisch" if str(x).strip().isdigit() else "Alphanumerisch"
        )

        # 父子标识
        df["is_parent"] = df["sku"] == df["parent_sku"]

        # 子体从父体继承 keywords & bullets
        parent_cols = [
            "generic_keywords", "bullet1", "bullet2",
            "bullet3", "bullet4", "bullet5",
        ]
        parent_df = (
            df[df["is_parent"]][["sku"] + parent_cols].set_index("sku")
        )
        parent_dict: dict = parent_df.to_dict(orient="index")

        def get_parent_data(row: pd.Series) -> pd.Series:
            if row["is_parent"]:
                return pd.Series([row[col] for col in parent_cols])
            p_data = parent_dict.get(row["parent_sku"], {})
            return pd.Series([p_data.get(col, "") for col in parent_cols])

        df[
            ["final_keywords"] + [f"final_bullet{i}" for i in range(1, 6)]
        ] = df.apply(get_parent_data, axis=1)

        return df

    # ==================== 模板填充 ====================

    def fill_data_to_template(self, ws: Worksheet, df: pd.DataFrame) -> None:
        """根据映射配置，将 DataFrame 逐行写入模板工作表。"""
        for r_idx, (_, row) in enumerate(df.iterrows(), start=self.START_ROW):
            # A. 一对一映射
            for src_col, tgt_idx in self._simple_idx.items():
                ws.cell(row=r_idx, column=tgt_idx, value=row[src_col])

            # B. 一对多映射
            for src_col, tgt_indices in self._multi_idx.items():
                for tgt_idx in tgt_indices:
                    ws.cell(row=r_idx, column=tgt_idx, value=row[src_col])

            # C. Bullet Points
            for i, tgt_idx in enumerate(self._bullet_idx):
                src_col = f"{self.BULLET_SRC_PREFIX}{i + 1}"
                ws.cell(row=r_idx, column=tgt_idx, value=row[src_col])

            # D. 图片
            images = [
                str(row[pos]).strip()
                for pos in self.IMG_SRC_COLS
                if str(row[pos]).strip()
            ]
            for img_i, url in enumerate(images[: self.MAX_IMAGES]):
                ws.cell(row=r_idx, column=self._img_start + img_i, value=url)
            for empty_i in range(len(images), self.MAX_IMAGES):
                ws.cell(row=r_idx, column=self._img_start + empty_i, value="")

            # E. 父子标识
            pc_val = "Parent" if row["is_parent"] else "Child"
            ws.cell(row=r_idx, column=self._pc_idx, value=pc_val)

    # ==================== 父体清除 ====================

    def clear_parent_rows(self, ws: Worksheet, df: pd.DataFrame) -> None:
        """抹平父体行的子体专属字段，并取消合并单元格。"""
        parent_local_indices = df[df["is_parent"]].index.tolist()
        if not parent_local_indices:
            return

        for local_row_idx in parent_local_indices:
            excel_row = self.START_ROW + local_row_idx

            # 取消该行所有合并单元格
            merged_to_remove = [
                mr
                for mr in ws.merged_cells.ranges
                if mr.min_row <= excel_row <= mr.max_row
            ]
            for mr in merged_to_remove:
                ws.unmerge_cells(str(mr))

            # 清空父体专属字段
            for col_idx in self.parent_clear_idx:
                ws.cell(row=excel_row, column=col_idx, value="")

            ws.cell(row=excel_row, column=self._pc_idx, value="Parent")

    # ==================== 拆分保存 ====================

    def save_split_workbooks(self, processed_df: pd.DataFrame) -> None:
        """将数据按 CHUNK_SIZE 拆分为多个文件并保存。"""
        total_rows = len(processed_df)
        if total_rows == 0:
            print("源数据为空，跳过。")
            return

        # 准备输出目录
        output_dir = self.OUTPUT_DIR
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"📁 创建输出目录: {output_dir}")

        base_name = os.path.splitext(os.path.basename(self.SOURCE_FILE))[0]
        num_chunks = math.ceil(total_rows / self.CHUNK_SIZE)

        print(f"\n数据预处理完成，共 {total_rows} 行。")
        print(f"按每文件 {self.CHUNK_SIZE} 行拆分，共 {num_chunks} 个文件。\n")

        for chunk_idx in range(num_chunks):
            start = chunk_idx * self.CHUNK_SIZE
            end = min(start + self.CHUNK_SIZE, total_rows)

            chunk_df = processed_df.iloc[start:end].reset_index(drop=True)
            chunk_rows = len(chunk_df)

            wb = load_workbook(self.TEMPLATE_FILE, keep_vba=False)
            if self.SHEET_NAME not in wb.sheetnames:
                wb.close()
                raise ProcessingError(
                    f"模板中未找到 '{self.SHEET_NAME}' 工作表"
                )
            ws = wb[self.SHEET_NAME]

            # 填充 & 清除
            self.fill_data_to_template(ws, chunk_df)
            self.clear_parent_rows(ws, chunk_df)

            # 删除多余空行
            last_used_row = self.START_ROW + chunk_rows - 1
            start_delete_row = last_used_row + 1
            if ws.max_row >= start_delete_row:
                ws.delete_rows(
                    start_delete_row, amount=ws.max_row - start_delete_row + 1
                )

            output_name = f"{base_name}{self.OUTPUT_SUFFIX}_part{chunk_idx + 1}.xlsx"
            if output_dir:
                output_path = os.path.join(output_dir, output_name)
            else:
                output_path = output_name

            print(
                f"  [{chunk_idx + 1}/{num_chunks}] 💾 {output_path}"
                f"  ({chunk_rows} 行)"
            )
            wb.save(output_path)
            wb.close()

        print(f"\n✅ 全部 {num_chunks} 个文件保存完毕。")

    # ==================== 入口 ====================

    def run(self) -> None:
        """执行完整处理流水线：读取 → 预处理 → 拆分写入。"""
        print(f"🚀 开始执行 Amazon 模板处理程序 [{self.__class__.__name__}]")

        if not self.SOURCE_FILE:
            raise ProcessingError("SOURCE_FILE 未设置")
        if not self.TEMPLATE_FILE:
            raise ProcessingError("TEMPLATE_FILE 未设置")

        print(f"\n[1/3] 读取源数据: {self.SOURCE_FILE}")
        if not os.path.exists(self.SOURCE_FILE):
            raise ProcessingError(f"源文件不存在: {self.SOURCE_FILE}")
        if not os.path.exists(self.TEMPLATE_FILE):
            raise ProcessingError(f"模板文件不存在: {self.TEMPLATE_FILE}")

        raw_df: pd.DataFrame = pd.read_excel(
            self.SOURCE_FILE, sheet_name=0, header=None, dtype=str
        )

        print("\n[2/3] 预处理源数据...")
        processed_df = self.preprocess_source_data(raw_df)

        print("\n[3/3] 执行拆分写入...")
        self.save_split_workbooks(processed_df)

        print("\n🎉 全部流程执行完毕！")

    def dry_run(self) -> None:
        """只做预处理，不写文件 — 用于验证配置。"""
        print(f"[DRY-RUN] 模式 [{self.__class__.__name__}]")

        if not os.path.exists(self.SOURCE_FILE):
            raise ProcessingError(f"源文件不存在: {self.SOURCE_FILE}")
        if not os.path.exists(self.TEMPLATE_FILE):
            raise ProcessingError(f"模板文件不存在: {self.TEMPLATE_FILE}")

        raw_df = pd.read_excel(
            self.SOURCE_FILE, sheet_name=0, header=None, dtype=str
        )
        processed_df = self.preprocess_source_data(raw_df)

        total_rows = len(processed_df)
        num_chunks = math.ceil(total_rows / self.CHUNK_SIZE)
        parent_count = processed_df["is_parent"].sum()
        child_count = total_rows - parent_count

        print(f"\n[SUMMARY] 预处理结果:")
        print(f"   总行数:     {total_rows}")
        print(f"   父体:       {parent_count}")
        print(f"   子体:       {child_count}")
        print(f"   拆分文件数: {num_chunks} (每 {self.CHUNK_SIZE} 行)")
        print(f"   列数:       {len(processed_df.columns)}")
        print(f"   输出目录:   {self.OUTPUT_DIR}")
        print(f"   模板文件:   {self.TEMPLATE_FILE}")
        print(f"\n[DONE] Dry-run 完成（未写入任何文件）。")
