# 品牌处理器重构方案

> 状态：⏳ 待审批 | 分析时间：2026-06-12

---

## 1. 现状：8 个文件的差异矩阵

### 1.1 100% 重复的代码块（每个文件都有，逐字相同）

| 代码块 | azd 行号 | hd 行号 | 其他文件 | 行数 |
|--------|----------|---------|---------|------|
| `preprocess_source_data()` | 97-137 | 115-155 | 完全相同 | 41 |
| `fill_data_to_template()` | 141-177 | 159-195 | 完全相同 | 37 |
| `clear_parent_rows()` | 181-204 | 199-222 | 完全相同 | 24 |
| `save_split_workbooks()` | 208-265 | 226-283 | 完全相同 | 58 |
| `main()` | 270-285 | 288-303 | 完全相同 | 16 |
| imports + 注释分隔线 | 1-8 | 1-8 | 完全相同 | 8 |
| Config 公共属性 | 19-43 | 19-43 | 完全相同 | 25 |
| **合计重复/每文件** | | | | **~209 行** |
| **8 文件总重复** | | | | **~1,672 行** |

### 1.2 每个品牌独有的配置项

| 配置项 | azd | azh | hd | hh | kle | std | td | th |
|--------|-----|-----|-----|-----|-----|-----|-----|-----|
| SOURCE_FILE | 610azd | 507azh | 610hd | 611hh | 609kle | 521std | 611WMtd | 609WM2 |
| CHUNK_SIZE | 500 | 500 | 1000 | 1000 | 1000 | 500 | 1000 | 500 |
| TEMPLATE 后缀 | azd | azh | hd (.xlsx!) | hh | kle | std | td | th |
| size_class 列 | AD | AD | **BA** | **BA** | AD | ~~注释掉~~ | AD | AD |
| final_keywords | CE ✅ | CE ✅ | CE ✅ | ~~注释掉~~ | CE ✅ | CE ✅ | CE ✅ | CE ✅ |

### 1.3 MULTI_MAP 中 size 列差异

| 品牌 | size 目标列 |
|------|------------|
| azd, azh, kle, td, th | `["AE", "CH", "FC"]` |
| hd, hh | `["BB", "CH", "FC"]` |
| std | `["CH", "FC"]` (缺 AE) |

### 1.4 PARENT_CLEAR_LETTERS 差异

| 品牌 | 额外清空列（相对于公共集） |
|------|--------------------------|
| azd, azh, kle | `V,W,X,AA,AB,AC,AD,AE,AF,AG,AH,BY,BZ,CF,CG,CH,FC` |
| hd, hh | 用 `BA,BB,BC,BF,BG,AZ` 替代 `AC,AD,AE,AF,AG,AH`，额外有 `AZ` |
| std | 最少：`V,W,AA,AB,BY,BZ,CF,CG,CH,FC` (无 X, AC~AH) |
| td, th | `V,W,X,AA,AB,AC,AD,AE,AG,AH,AZ,BY,BZ,CF,CG,CH,FC` (无 AF) |

### 1.5 SIZE_REPLACEMENTS 差异

- **th 独有**：额外 4 条规则 `L2→XXL, L3→3XL, L4→4XL, L5→5XL`
- 其余 7 个品牌：完全相同（9 条 XL 规则 + "one size"）

---

## 2. 重构方案设计

### 2.1 目标目录结构

```
autoUpload/
├── base_processor.py          # 公共基类 BaseProcessor
├── brands/                    # 品牌配置（每个 ~15 行）
│   ├── __init__.py
│   ├── azd.py
│   ├── azh.py
│   ├── hd.py
│   ├── hh.py
│   ├── kle.py
│   ├── std.py
│   ├── td.py
│   └── th.py
├── run_brand.py               # 统一 CLI 入口
├── process.py                 # 不动
├── split_excel.py             # 不动
├── prov_uploadEcel_Split_*.py # 保留为 import 原文件的兼容入口
└── temp/                      # 不动
```

### 2.2 基类 `BaseProcessor` 设计

```python
# base_processor.py

import pandas as pd
import re, math, os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


class BaseProcessor:
    # ========== 子类必须覆盖的配置 ==========
    SOURCE_FILE: str = ""
    TEMPLATE_FILE: str = ""
    CHUNK_SIZE: int = 500
    SIZE_CLASS_COL: str = "AD"           # None 表示不写 size_class
    INCLUDE_FINAL_KEYWORDS: bool = True   # 是否写 final_keywords

    # 公共列映射（子类可覆盖）
    @property
    def SIMPLE_MAP(self):
        m = {
            "sku": "B", "parent_sku": "BY", "title": "G",
            "standprice": "V", "package_weight": "GD",
            "list_price_tax": "KP", "final_keywords": "CE",
        }
        if self.SIZE_CLASS_COL:
            m["size_class"] = self.SIZE_CLASS_COL
        if not self.INCLUDE_FINAL_KEYWORDS:
            del m["final_keywords"]
        return m

    @property
    def MULTI_MAP(self):
        return {
            "color": ["CF", "CG"],
            "size":  self.SIZE_TARGET_COLS,
        }

    # 子类覆盖
    SIZE_TARGET_COLS = ["AE", "CH", "FC"]
    PARENT_CLEAR_LETTERS = [
        "V","W","X","AA","AB","AC","AD","AE","AF","AG","AH",
        "BY","BZ","CF","CG","CH","FC",
    ]
    SIZE_REPLACEMENTS = {
        r"\bXXXL\b":"3XL", r"\bXXXXL\b":"4XL", r"\bXXXXXL\b":"5XL",
        r"\bXXXXXXL\b":"6XL", r"\bXXXXXXXL\b":"7XL",
        r"\bXXXXXXXXL\b":"8XL", r"\bXXXXXXXXXL\b":"9XL",
        r"\bone size\b":"Einheitsgröße",
    }

    # ========== 公共常量 ==========
    OUTPUT_SUFFIX = "_prov"
    MAX_SOURCE_COLS = 60
    START_ROW = 4
    MAX_IMAGES = 7
    SRC = { ... }        # 不变的源列索引
    BULLET_SRC_PREFIX = "final_bullet"
    BULLET_TGT_COLS = ["CO","CP","CQ","CR","CS"]
    IMG_SRC_COLS = [10,11,12,13,14,15,16]
    IMG_TGT_START = "BH"
    PC_COL = "BX"
    OUTPUT_DIR = "./prov_output"
    TITLE_PREFIX = ""
    SHEET_NAME = "Vorlage"

    # ========== 公共方法 ==========
    @property
    def parent_clear_idx(self):  # 延迟计算
        return [column_index_from_string(c) for c in self.PARENT_CLEAR_LETTERS]

    def preprocess_source_data(self, df: pd.DataFrame) -> pd.DataFrame:
        ...  # 从现有代码搬来，self.Config.xxx → self.xxx

    def fill_data_to_template(self, ws, df: pd.DataFrame):
        ...

    def clear_parent_rows(self, ws, df: pd.DataFrame):
        ...

    def save_split_workbooks(self, processed_df: pd.DataFrame):
        ...

    def run(self):
        """统一入口"""
        print(f"🚀 开始处理品牌: {self.__class__.__name__}")
        raw_df = pd.read_excel(self.SOURCE_FILE, sheet_name=0, header=None, dtype=str)
        processed_df = self.preprocess_source_data(raw_df)
        self.save_split_workbooks(processed_df)
        print("🎉 完成！")
```

### 2.3 子类示例（azd）

```python
# brands/azd.py
from base_processor import BaseProcessor

class AzdProcessor(BaseProcessor):
    SOURCE_FILE = "610azd.xlsx"
    TEMPLATE_FILE = "./temp/prov_azdTemp.xlsm"
    CHUNK_SIZE = 500
    # 其余全部继承默认值
```

### 2.4 各子类需覆盖的属性

| 品牌文件 | 需覆盖的属性 | 覆盖原因 |
|----------|-------------|----------|
| azd, azh, kle | 仅 SOURCE_FILE, TEMPLATE_FILE | 全默认 |
| hd | SOURCE_FILE, TEMPLATE_FILE, CHUNK_SIZE=1000, SIZE_CLASS_COL="BA", SIZE_TARGET_COLS=["BB","CH","FC"], PARENT_CLEAR_LETTERS (改) | 列结构不同 |
| hh | 同 hd + INCLUDE_FINAL_KEYWORDS=False | keywords 不输出 |
| std | SOURCE_FILE, TEMPLATE_FILE, SIZE_CLASS_COL=None, SIZE_TARGET_COLS=["CH","FC"], PARENT_CLEAR_LETTERS (少列) | 缺 size_class, 缺 AE |
| td | SOURCE_FILE, TEMPLATE_FILE, CHUNK_SIZE=1000, PARENT_CLEAR_LETTERS (加 AZ, 缺 AF) | |
| th | SOURCE_FILE, TEMPLATE_FILE, PARENT_CLEAR_LETTERS (加 AZ, 缺 AF), SIZE_REPLACEMENTS (加 L2~L5) | 尺码规则不同 |

### 2.5 CLI 入口设计

```bash
# 方式 1：通过品牌代码运行
python run_brand.py --brand azd
python run_brand.py --brand hd --input 610hd.xlsx --chunk 800

# 方式 2：兼容旧调用方式（原文件仍可独立运行）
python prov_uploadEcel_Split_azd.py
# → 内部 from brands.azd import AzdProcessor; AzdProcessor().run()
```

`run_brand.py` 参数：
| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--brand` | 品牌代码 (azd/azh/hd/hh/kle/std/td/th) | **必填** |
| `--input` | 源 Excel 文件（覆盖 Config） | Config 默认值 |
| `--template` | 模板文件路径（覆盖 Config） | Config 默认值 |
| `--chunk` | 分块行数 | Config 默认值 |
| `--output-dir` | 输出目录 | `./prov_output` |

---

## 3. 兼容性策略

### 3.1 旧文件保留为 stub

8 个原始 `prov_uploadEcel_Split_*.py` 改为 2 行桩代码：

```python
# prov_uploadEcel_Split_azd.py（重构后）
from brands.azd import AzdProcessor
if __name__ == "__main__":
    AzdProcessor().run()
```

### 3.2 `build_all.bat` 不需要改

因为入口文件名未变。

---

## 4. 风险点与回滚

| 风险 | 缓解 |
|------|------|
| 属性名变更导致行为差异 | 逐文件 diff 测试：重构后输出文件与原文件逐字节对比 |
| SIZE_REPLACEMENTS 是 dict，直接赋值会共享引用 | 基类用 `@property` 返回新 dict，或子类用 `deepcopy` |
| openpyxl `keep_vba` 参数（原为 False） | 确认 False 是正确行为 |
| th 的额外 L2~L5 规则 | 子类覆盖 `SIZE_REPLACEMENTS` 属性 |

**回滚方案**：`git revert` 重构 commit，8 个原始文件完整保留在 git 历史中。

---

## 5. 执行步骤

| 步骤 | 内容 | 验证方式 |
|------|------|----------|
| 1 | 创建 `base_processor.py` | 语法检查 |
| 2 | 创建 `brands/` 目录 + 8 个子类文件 | 语法检查 |
| 3 | 修改 8 个原文件为 stub | 回归测试 |
| 4 | 创建 `run_brand.py` CLI 入口 | 参数解析测试 |
| 5 | 回归测试：逐个品牌运行，对比输出文件 | diff 输出目录 |
| 6 | git commit，标记重构完成 | — |

---

> ⚠️ **本方案尚未执行。请审阅后确认是否继续。**
