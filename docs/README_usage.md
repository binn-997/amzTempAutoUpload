# autoUpload — 完整使用手册

> **版本**: v2.1 | **更新日期**: 2026-07-01 | **Python**: 3.10+

---

## 目录

1. [项目概述](#1-项目概述)
2. [环境准备](#2-环境准备)
3. [项目结构](#3-项目结构)
4. [配置文件详解](#4-配置文件详解)
5. [CLI 命令行用法](#5-cli-命令行用法)
6. [兼容入口脚本](#6-兼容入口脚本)
7. [独立工具](#7-独立工具)
8. [常见操作场景](#8-常见操作场景)
9. [添加新产品类型](#9-添加新产品类型)
10. [开发者指南](#10-开发者指南)
11. [故障排除](#11-故障排除)

---

## 1. 项目概述

autoUpload 是一套用于 **Amazon 德国站 Listing 批量上传** 的数据处理工具集。它将商品源 Excel 数据（SKU、颜色、尺码、价格、图片 URL、Bullet Points 等）经过预处理后，按 Amazon 库存文件模板格式拆分为可分批上传的 Excel 文件。

### 核心功能

| 功能 | 入口 | 说明 |
|------|------|------|
| **品牌数据处理** | `src/main.py process` | 读取源数据 → 数据清洗 → 模板填充 → 分块保存 |
| **查看产品类型** | `src/main.py list-types` | 列出配置文件中所有可用的产品类型 |
| **颜色翻译** | `process.py` | 将英文颜色名翻译为德文，并标注变体 V1/V2 后缀 |
| **Excel 拆分** | `split_excel.py` | 将大 Excel 按指定行数拆分为多个小文件 |
| **利润计算器** | `src/profit_calculator.py` | 多站点（US/CA/DE）利润核算，含扣点率拆解 |
| **利润定价** | `src/main.py price` | 根据成本+重量反算销售价，回填到源文件 AM 列 |
| **颜色翻译** | `src/main.py translate` | 颜色英→德翻译 + 变体 V1/V2 标注，原地修改源文件 |

### 8 个产品类型

| 标识 | 描述 | 兼容入口脚本 |
|------|------|-------------|
| `azd` | Anzug Damen（女士西装） | `prov_uploadEcel_Split_azd.py` |
| `azh` | Anzug Herren（男士西装） | `prov_uploadEcel_Split_azh.py` |
| `hd` | Hose Damen（女士裤装） | `prov_uploadEcel_Split_hd.py` |
| `hh` | Hose Herren（男士裤装） | `prov_uploadEcel_Split_hh.py` |
| `kle` | Kleid（连衣裙） | `prov_uploadEcel_Split_kle.py` |
| `std` | Strumpfhose Damen（连裤袜） | `prov_uploadEcel_Split_std.py` |
| `td` | Top Damen（女士上装） | `prov_uploadEcel_Split_td.py` |
| `th` | Top Herren（男士上装） | `prov_uploadEcel_Split_th.py` |

---

## 2. 环境准备

### 2.1 安装 Python

需要 **Python 3.10 或更高版本**。下载地址：<https://www.python.org/downloads/>

安装时勾选 "Add Python to PATH"。

### 2.2 安装依赖

```bash
cd autoUpload

# 安装运行时依赖（必须）
pip install -r requirements.txt

# 安装开发依赖（可选：类型检查、测试、代码格式化）
pip install -r requirements-dev.txt
```

### 2.3 验证安装

```bash
python -c "import pandas; import openpyxl; print('OK')"
```

### 2.4 依赖清单

#### 运行时依赖 (`requirements.txt`)

| 包名 | 版本 | 用途 | 必须 |
|------|------|------|:--:|
| `pandas` | ≥2.0, <4.0 | Excel 读取与数据清洗 | ✅ |
| `openpyxl` | ≥3.1, <4.0 | Excel 模板读写 | ✅ |
| `pyyaml` | ≥6.0, <7.0 | YAML 配置文件解析 | 可选 |
| `pyinstaller` | ≥6.0, <7.0 | 打包为 .exe | 可选 |

> **说明**：如果不使用 YAML 配置文件，可将配置存为 JSON 格式，无需安装 `pyyaml`。详见 [配置文件详解](#4-配置文件详解)。

#### 开发依赖 (`requirements-dev.txt`)

| 包名 | 用途 |
|------|------|
| `mypy` | 静态类型检查 |
| `ruff` | 代码格式化 & Lint |
| `pytest` + `pytest-cov` | 单元测试 & 覆盖率 |
| `pyinstaller` | 打包为 .exe |

---

## 3. 项目结构

```
autoUpload/
├── README.md                          # 项目入口说明
├── CLAUDE.md                          # 开发者指引
├── requirements.txt                   # Python 运行时依赖
├── requirements-dev.txt               # Python 开发依赖
├── .gitignore                         # Git 忽略规则
│
├── config/
│   └── categories.yaml                # ★ 产品类型配置（核心文件，运营人员编辑）
│
├── src/                               # 核心 Python 包
│   ├── __init__.py
│   ├── main.py                        # ★ 统一 CLI 入口
│   ├── config_loader.py               # 配置加载 & 合并工具
│   ├── profit_calculator.py           # ★ 多站点利润计算器
│   ├── color_translator.py            # ★ 颜色翻译 + 变体检测
│   └── processors/
│       ├── __init__.py
│       └── base_processor.py          # ★ 通用处理器基类（全部处理逻辑）
│
├── temp/                              # Excel 模板文件（8 个品牌）
│   ├── prov_azdTemp.xlsm
│   ├── prov_azhTemp.xlsm
│   ├── prov_hdTemp.xlsx               # ← hd 模板为 .xlsx 格式
│   ├── prov_hhTemp.xlsm
│   ├── prov_kleTemp.xlsm
│   ├── prov_stdTemp.xlsm
│   ├── prov_tdTemp.xlsm
│   └── prov_thTemp.xlsm
│
├── prov_uploadEcel_Split_azd.py       # 兼容入口脚本 (×8)
├── prov_uploadEcel_Split_azh.py
├── prov_uploadEcel_Split_hd.py
├── prov_uploadEcel_Split_hh.py
├── prov_uploadEcel_Split_kle.py
├── prov_uploadEcel_Split_std.py
├── prov_uploadEcel_Split_td.py
├── prov_uploadEcel_Split_th.py
│
├── process.py                         # 独立工具：颜色翻译
├── split_excel.py                     # 独立工具：Excel 拆分
│
├── data/                              # 源数据 Excel 文件 (gitignored)
├── prov_output/                       # 处理输出目录 (gitignored)
├── prov已上传/                         # 已上传文件归档 (gitignored)
├── docs/                              # 文档归档
│   ├── CHANGELOG.md
│   ├── README_usage.md                # ← 本文档
│   └── ...
├── dist/                              # PyInstaller 打包输出
└── build/                             # PyInstaller 构建中间产物
```

---

## 4. 配置文件详解

配置文件 `config/categories.yaml` 是整个系统的核心，所有产品类型的差异化参数均在此定义。

### 4.1 文件格式

默认使用 **YAML** 格式。也支持 **JSON**（将扩展名改为 `.json`，内容使用 JSON 语法即可）。

推荐使用 YAML，因为：
- 支持注释（JSON 不支持）
- 语法简洁，无需引号和大括号
- 运营人员更易阅读

### 4.2 配置结构总览

```yaml
defaults:           # 全局默认值（所有产品类型共用）
  chunk_size: 500
  title_prefix: ""
  # ...

source_columns:     # 源数据列索引（0-based），所有产品类型共用
  sku: 1
  parent_sku: 2
  # ...

template_columns:   # 模板目标列字母，所有产品类型共用
  simple:
    sku: "B"
  multi:
    color: ["CF", "CG"]
  # ...

categories:         # ★ 产品类型定义（只需定义与 defaults 不同的部分）
  azd:
    description: "..."
    default_source: "..."
    template_file: "..."
    chunk_size: 1000
```

### 4.3 `defaults` — 全局默认值

所有产品类型的公共默认参数。

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `chunk_size` | int | `500` | 每个输出文件的最大数据行数 |
| `title_prefix` | str | `""` | 加在商品标题前的文本 |
| `output_dir` | str | `"./prov_output"` | 输出文件存放目录 |
| `output_suffix` | str | `"_prov"` | 输出文件名后缀 |
| `include_final_keywords` | bool | `true` | 是否写入 `final_keywords` 列 |
| `size_class_col` | str 或 null | `"AD"` | `size_class` 写入列，`null` 不输出 |
| `size_target_cols` | list[str] | `["AE","CH","FC"]` | `size` 字段一对多映射目标列 |
| `parent_clear_letters` | list[str] | （17 列） | 父体行需清空的列字母 |
| `size_replacements` | dict | （8 条规则） | 尺码正则替换（正则 → 替换文本） |

**合并规则**：

| 值类型 | 合并方式 | 示例 |
|--------|----------|------|
| 标量（str/int/bool） | 产品类型覆盖 defaults | `chunk_size: 1000` 覆盖 `500` |
| 字典（dict） | **深度合并**，产品类型键优先 | `size_replacements` 追加而不是替换 |
| 列表（list） | **完全替换** | `parent_clear_letters` 直接替换 |

### 4.4 `source_columns` — 源数据列索引

定义源 Excel 中各字段所在的列号（**0-based**：A=0, B=1, ...）。

| 字段 | 默认列号 | 说明 |
|------|----------|------|
| `sku` | 1 | 商品 SKU |
| `parent_sku` | 2 | 父体 SKU |
| `title` | 3 | 商品标题 |
| `color` | 6 | 颜色 |
| `size` | 7 | 尺码 |
| `package_weight` | 9 | 包装重量 |
| `standprice` | 38 | 标准售价 |
| `list_price_tax` | 39 | 含税标价 |
| `generic_keywords` | 45 | 通用关键词 |
| `bullet1` ~ `bullet5` | 40~44 | 5 个 Bullet Points |
| `images` | [10~16] | 7 张商品图片 URL |

> 如源 Excel 列结构变化，在此处调整对应列号即可。

### 4.5 `template_columns` — 模板目标列

定义模板 Excel 中各字段对应的列字母。

| 子节点 | 说明 |
|--------|------|
| `simple` | 一对一映射（源字段名 → 模板列字母） |
| `multi` | 一对多映射（源字段名 → [模板列字母列表]） |
| `bullets` | 5 个 Bullet Points 目标列 |
| `images_start` | 7 张图片起始列 |
| `parent_child` | 父子标识列（Parent/Child） |

> 如 Amazon 模板列结构变化，在此处调整即可。

### 4.6 `categories` — 产品类型定义

每个产品类型一个条目。

**必填字段**：

| 字段 | 说明 | 示例 |
|------|------|------|
| `description` | 产品类型描述 | `"Anzug Damen（女士西装）"` |
| `default_source` | 默认源 Excel 路径 | `"data/624madchen.xlsx"` |
| `template_file` | 模板文件路径 | `"./temp/prov_azdTemp.xlsm"` |

**可选字段**（未指定则继承 `defaults`）：

`chunk_size`、`title_prefix`、`output_dir`、`include_final_keywords`、`size_class_col`、`size_target_cols`、`parent_clear_letters`、`size_replacements`

### 4.7 各产品类型差异一览

| 参数 | azd/azh/kle | hd | hh | std | td | th |
|------|:--:|:--:|:--:|:--:|:--:|:--:|
| `chunk_size` | 1000 | 1000 | 1000 | 500 | 3000 | 1200 |
| `size_class_col` | AD | BA | BA | null | AD | AD |
| `include_final_keywords` | true | true | **false** | true | true | true |
| `size_target_cols` | AE/CH/FC | BB/CH/FC | BB/CH/FC | CH/FC | AE/CH/FC | AE/CH/FC |
| `parent_clear_letters` | 17列 | 17列(异) | 17列(异) | 10列 | 17列(异) | 17列(异) |
| 额外尺码规则 | — | — | — | — | — | L2→XXL 等 4 条 |

---

## 5. CLI 命令行用法

### 5.1 查看帮助

```bash
python -m src.main --help
python -m src.main process --help
python -m src.main list-types --help
python -m src.main price --help
```

### 5.2 列出可用产品类型

```bash
python -m src.main list-types
```

输出：

```
📋 配置文件: D:\...\config\categories.yaml

标识       描述
--------------------------------------------------
azd      Anzug Damen（女士西装）
azh      Anzug Herren（男士西装）
hd       Hose Damen（女士裤装）
hh       Hose Herren（男士裤装）
kle      Kleid（连衣裙）
std      Strumpfhose Damen（女士连裤袜）
td       Top Damen（女士上装）
th       Top Herren（男士上装）
```

### 5.3 处理产品类型

**基本用法**（使用配置文件中的默认源文件和参数）：

```bash
python -m src.main process --type azd
```

**指定源文件**（覆盖配置默认值）：

```bash
python -m src.main process --type azd --source data/新数据.xlsx
```

**覆盖分块大小**：

```bash
python -m src.main process --type td --chunk 500
```

**指定自定义模板**：

```bash
python -m src.main process --type hd --template temp/my_template.xlsm
```

**自定义输出目录**：

```bash
python -m src.main process --type th --output-dir ./my_output
```

**使用 JSON 配置文件**：

```bash
python -m src.main process --type azd --config config/categories.json
```

**Dry-run 模式**（预览分析，不写文件）：

```bash
python -m src.main process --type th --dry-run
```

输出示例：

```
[DRY-RUN] 模式 [BaseProcessor]
📋 配置文件: .../config/categories.yaml
🏷️  产品类型: th (Top Herren（男士上装）)
预处理源数据：清洗空值、替换尺码、继承父体数据...

[SUMMARY] 预处理结果:
   总行数:     320
   父体:       15
   子体:       305
   拆分文件数: 1 (每 1200 行)
   列数:       68
   输出目录:   ./prov_output
   模板文件:   ./temp/prov_thTemp.xlsm

[DONE] Dry-run 完成（未写入任何文件）。
```

### 5.4 参数速查

| 参数 | 必选 | 类型 | 说明 | 默认值 |
|------|:--:|------|------|--------|
| `--type` | ✅ | str | 产品类型标识 | — |
| `--source` | ❌ | str | 源 Excel 文件路径 | 配置文件 `default_source` |
| `--template` | ❌ | str | 模板文件路径 | 配置文件 `template_file` |
| `--chunk` | ❌ | int | 分块行数 | 配置文件 `chunk_size` |
| `--output-dir` | ❌ | str | 输出目录 | 配置文件 `output_dir` |
| `--config` | ❌ | str | 配置文件路径 | `config/categories.yaml` |
| `--dry-run` | ❌ | flag | 只预览分析，不写文件 | false |

### 5.5 两种运行方式

```bash
# 方式 A：作为模块运行（推荐，适用于所有环境）
python -m src.main process --type azd

# 方式 B：直接运行脚本（需在项目根目录下）
python src/main.py process --type azd
```

### 5.6 利润定价（price）

根据商品的成本价和重量，反算出达到目标毛利率所需的最低销售价，并**直接回填到源文件**的指定列。

**基本用法**：

```bash
# 德国站（默认），目标毛利率 25%，从 I 列读成本、J 列读重量，写入 AM 列
python -m src.main price --source data/630td.xlsx

# 自定义目标毛利率和汇率
python -m src.main price --source data/630td.xlsx --target 0.30 --rate 8.0

# 美国站
python -m src.main price --source data/xxx.xlsx --station US --rate 7.0

# 自定义列映射
python -m src.main price --source data/xxx.xlsx --cost-col E --weight-col F --price-col Z
```

**参数速查**：

| 参数 | 必选 | 类型 | 默认值 | 说明 |
|------|:--:|------|--------|------|
| `--source` | ✅ | str | — | 源 Excel 文件路径（结果直接回填到该文件） |
| `--target` | ❌ | float | `0.25` | 目标毛利率（0.25 = 25%） |
| `--rate` | ❌ | float | `7.8` | 汇率 RMB/外币 |
| `--shipping-price` | ❌ | float | `0` | 向买家收取的运费（外币），0 表示包邮模式 |
| `--shipping-method` | ❌ | str | `ceil` | 运费查表方式：`ceil`（向上取档）\| `linear`（线性插值）\| `floor`（向下取档） |
| `--station` | ❌ | str | `DE` | 目标站点：`DE` \| `US` \| `CA` |
| `--cost-col` | ❌ | str | `I` | 成本价（RMB）所在列字母 |
| `--weight-col` | ❌ | str | `J` | 重量（g）所在列字母 |
| `--price-col` | ❌ | str | `AM` | 销售价写入列字母 |

**运行示例输出**：

```
📋 源文件:   data/630td.xlsx
🏷️  站点:     DE
🎯 目标毛利: 25.0%
💱 汇率:     7.8
📦 运费查表: ceil
📐 成本列:   I  重量列: J  →  写入列: AM

✅ 有效行数: 1199  (0 行跳过)
💰 价格区间: EUR 17.44 ~ 27.53
📊 价格均值: EUR 19.44  中位数: EUR 19.34

💾 回填到源文件: data/630td.xlsx
🎉 已将 1199 条销售价写入 AM 列。
```

**工作原理**：

1. 读取源 Excel（pandas，无表头）
2. 逐行从 `--cost-col` 取成本、`--weight-col` 取重量
3. 跳过空行或无效数据（成本/重量 ≤0）
4. 调用 `ProfitCalculator.breakeven_price()` 反算最低售价（自动迭代匹配正确的扣点率档位）
5. 用 `openpyxl` 将价格**回填到源文件**的 `--price-col` 列（带表头"销售价(EUR)"）
6. 其他列及格式保持不变

> 扣点率档位及公式原理详见 [docs/profit_calculator.md](profit_calculator.md)。

### 5.7 颜色翻译（translate）

对源文件中的颜色列执行：**英文→德文翻译** + **同款变体 V1/V2 后缀标注**，并**直接回填到源文件**。

**基本用法**：

```bash
# 翻译颜色 + 变体检测（默认行为）
python -m src.main translate --source data/701th.xlsx

# 仅翻译颜色，不检测变体
python -m src.main translate --source data/701th.xlsx --no-variant

# 仅检测变体，不翻译
python -m src.main translate --source data/701th.xlsx --no-translate

# 自定义颜色/尺码列
python -m src.main translate --source data/xxx.xlsx --color-col E --size-col F

# 指定工作表索引
python -m src.main translate --source data/xxx.xlsx --sheet 0
```

**参数速查**：

| 参数 | 必选 | 类型 | 默认值 | 说明 |
|------|:--:|------|--------|------|
| `--source` | ✅ | str | — | 源 Excel 文件路径（结果直接回填） |
| `--color-col` | ❌ | str | `G` | 颜色列字母 |
| `--size-col` | ❌ | str | `H` | 尺码列字母（用于变体签名匹配，不修改） |
| `--sheet` | ❌ | int | `0` | 工作表索引（0 = 第一个） |
| `--no-translate` | ❌ | flag | false | 跳过颜色翻译环节 |
| `--no-variant` | ❌ | flag | false | 跳过变体检测环节 |

**运行示例输出**：

```
📋 源文件:   data/701th.xlsx
🎨 颜色列:   G (0-based 索引 6)
📏 尺码列:   H (0-based 索引 7)
📄 工作表:   第 0 个
🌐 颜色翻译: ✅ 开
🔢 变体检测: ✅ 开

📊 总处理行数:   4
📊 颜色翻译数:   1 个颜色组
📊 变体标注数:   0 行
📊 实际修改行数: 1

💾 回填到源文件: data/701th.xlsx
🎉 已将 1 行颜色值写回 G 列。
```

**处理逻辑**：

1. 读取源 Excel（pandas，无表头）
2. 按 `parent_sku` 列（C 列）将数据分组为产品变体块（替代旧版空行分割方式）
3. 每个变体块内：将连续相同颜色的子体行聚合为「颜色组」
4. 颜色翻译：英文 → 德文，支持数字后缀剥离（如 `Blue2` → `Blau2`）
5. 变体检测：签名 = `(翻译后颜色, 尺码元组)`。同签名在块内重复出现时，依次标注 `V1`、`V2`、`V3`...
6. 用 openpyxl 原地回填颜色列，**其他列、格式、工作表均保持不变**

**内置颜色映射**（45 条，英 → 德）：

| 英文 | 德文 | 英文 | 德文 |
|------|------|------|------|
| Red | Rot | Black | Schwarz |
| Blue | Blau | White | Weiß |
| Green | Grün | Yellow | Gelb |
| Pink | Rosa | Grey | Grau |
| Purple | Lila | Brown | Braun |
| Orange | Orange | Gold | Gold |
| ... | ... | ... | ... |

> 完整列表见 `src/color_translator.py` 中 `COLOR_TRANSLATIONS` 字典。新增颜色直接编辑该字典即可。

**与旧版 `process.py` 的区别**：

| | `process.py`（旧） | `src/main.py translate`（新） |
|------|------|------|
| 数据格式 | 2 列（颜色, 尺码） | 全列源数据文件（~47 列） |
| 分组方式 | 空行分隔区块 | `parent_sku` 列（C 列） |
| 写入方式 | 新文件（时间戳命名） | 原地修改源文件 |
| 颜色列 | A 列 | 可配置（默认 G 列） |
| 尺码列 | B 列 | 可配置（默认 H 列） |
| 入口 | `python process.py` | `python -m src.main translate` |
| 多工作表 | 全部写回 | 仅指定工作表 |

---

## 6. 兼容入口脚本

8 个 `prov_uploadEcel_Split_*.py` 文件是**向后兼容入口**，用于习惯旧版操作的用户。

### 6.1 用法

```bash
python prov_uploadEcel_Split_azd.py   # 等价于 python -m src.main process --type azd
python prov_uploadEcel_Split_th.py    # 等价于 python -m src.main process --type th
```

### 6.2 工作原理

每个脚本内部：
1. 从 `config/categories.yaml` 加载对应产品类型的默认配置
2. 创建 `BaseProcessor` 实例
3. 调用 `.run()` 执行完整流水线

### 6.3 为新产品类型创建兼容入口

复制任一现有脚本，修改文件中的 `_TYPE_KEY` 变量：

```python
# 例如创建 prov_uploadEcel_Split_abc.py
_TYPE_KEY = "abc"   # 改为新产品类型标识
```

---

## 7. 独立工具

### 7.1 `process.py` — 颜色翻译

将源 Excel 中的英文颜色名翻译为德文，并自动检测变体（相同颜色+尺码组合但不同行），添加 V1/V2 后缀。

**用法**：

```bash
python process.py
```

**配置**（编辑 `process.py` 文件内部）：

```python
INPUT_FILE = "data/1.xlsx"
OUTPUT_FILE = f"data/processed_1_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
SHEET_NAME = "Sheet1"
```

**处理逻辑**：

1. 按空行分割数据区块
2. 区块内按连续相同颜色聚合为「颜色组」
3. 英文颜色名 → 德文翻译（42 条映射 + 数字后缀剥离）
4. 签名（颜色, 尺码元组）相同的组 → 标记 V1, V2, V3...

**内置颜色映射**（42 条）：

| 英文 | 德文 |
|------|------|
| Red | Rot |
| Black | Schwarz |
| Blue | Blau |
| White | Weiß |
| ... | ... |

> 完整列表见 `process.py` 中 `COLOR_TRANSLATIONS` 字典。新增颜色直接编辑该字典。

### 7.2 `split_excel.py` — Excel 拆分

将大 Excel 按指定行数上限拆分为多个小文件。

**用法**：

```bash
# 命令行
python split_excel.py -s data/大文件.xlsx -n 500 -o ./split_output

# Windows 拖放：将 Excel 拖到 exe 图标上即可
```

**参数**：

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-s` / `--source` | 源 Excel 文件 | 自动检测当前目录 |
| `-n` / `--chunk-size` | 每文件最大行数 | 500 |
| `-o` / `--output-dir` | 输出目录 | `./split_output` |
| `--header-rows` | 表头行数 | 3 |
| `--sheet` | 工作表名 | Vorlage |
| `--suffix` | 输出文件名后缀 | `_part` |

### 7.3 `src/profit_calculator.py` — 利润计算器

多站点（US / CA / DE）跨境电商利润计算模块。详见 [docs/profit_calculator.md](profit_calculator.md)。

---

## 8. 常见操作场景

### 场景 1：日常处理

```bash
python -m src.main process --type azd
# 输出在 prov_output/
```

### 场景 2：临时使用不同源文件

```bash
python -m src.main process --type hd --source data/新到货_hd.xlsx
```

### 场景 3：首次运行前预览

```bash
python -m src.main process --type th --dry-run
# 确认行数、父子体数量、拆分文件数无误后再正式运行
```

### 场景 4：永久修改某产品类型的默认源文件

编辑 `config/categories.yaml`：

```yaml
azd:
  default_source: "data/最新数据.xlsx"   # ← 改这里
```

之后每次 `--type azd` 都使用新文件，无需每次传 `--source`。

### 场景 5：调整分块大小

```bash
# 临时
python -m src.main process --type td --chunk 500

# 永久：编辑 config/categories.yaml 中 td 的 chunk_size
```

### 场景 6：拆分超大 Excel 后再处理

```bash
# 1. 先拆成 3000 行一份的小文件
python split_excel.py -s data/超大文件.xlsx -n 3000

# 2. 用品牌处理器处理第一份
python -m src.main process --type hd --source split_output/超大文件_part1.xlsx
```

### 场景 7：查看所有产品类型及其配置状态

```bash
python -m src.main list-types    # 列出类型
cat config/categories.yaml       # 查看详细配置
```

### 场景 8：颜色翻译 → 品牌处理流水线

```bash
# 方案 A：新 CLI 命令（推荐，原地修改源文件）
python -m src.main translate --source data/701th.xlsx
python -m src.main process --type th --source data/701th.xlsx

# 方案 B：旧版独立工具（2 列格式，生成新文件）
#   (修改 process.py 中 INPUT_FILE 指向源文件)
python process.py
python -m src.main process --type azd --source data/processed_1_20260626_120000.xlsx
```

---

## 9. 添加新产品类型

假设添加产品类型 `abc`，只需 **2 步**，无需编写 Python 代码。

### 步骤 1：编辑配置文件

在 `config/categories.yaml` 的 `categories` 节下添加：

```yaml
categories:
  # ... 已有 8 个条目 ...

  abc:
    description: "ABC 产品类型名称"
    default_source: "data/abc_源数据.xlsx"
    template_file: "./temp/prov_abcTemp.xlsm"
    chunk_size: 800

    # 如果列映射不同于默认值，添加对应字段：
    # size_class_col: "BA"
    # size_target_cols: ["BB", "CH", "FC"]
    # parent_clear_letters:
    #   - "V"
    #   - "W"
    #   - ...
    # size_replacements:         # 额外尺码规则（会与 defaults 合并）
    #   '\bL2\b': 'XXL'
```

### 步骤 2：放置模板文件

将 Amazon 库存文件模板放入 `temp/` 目录，命名为 `prov_abcTemp.xlsm`。

### 验证

```bash
python -m src.main list-types            # 应显示 abc
python -m src.main process --type abc --dry-run
python -m src.main process --type abc    # 正式运行
```

### （可选）创建兼容入口

```bash
# 复制任一现有脚本
cp prov_uploadEcel_Split_azd.py prov_uploadEcel_Split_abc.py
# 编辑文件，将 _TYPE_KEY = "azd" 改为 _TYPE_KEY = "abc"
```

---

## 10. 开发者指南

### 10.1 架构图

```
config/categories.yaml          ← 单一配置源
       │
       ▼
src/config_loader.py            ← 加载 & 合并配置
       │
       ▼
src/main.py                     ← CLI 入口
       │
       ▼
src/processors/base_processor.py ← 通用处理器（配置驱动）
       │
       ├── preprocess_source_data()  ← 数据清洗 & 父子继承
       ├── fill_data_to_template()   ← 逐行写入模板
       ├── clear_parent_rows()       ← 父体行清空
       └── save_split_workbooks()    ← 分块保存（BytesIO 优化）
```

### 10.2 钩子方法

`BaseProcessor` 提供两个钩子供特殊产品类型扩展：

```python
class BaseProcessor:
    def post_preprocess(self, df: pd.DataFrame) -> pd.DataFrame:
        """预处理后钩子。默认空操作。"""
        return df

    def post_fill_row(self, ws, excel_row, row_data) -> None:
        """单行填充后钩子。默认空操作。"""
        pass
```

### 10.3 异常体系

| 异常类 | 父类 | 用途 |
|--------|------|------|
| `ConfigError` | `Exception` | 配置文件加载/解析错误 |
| `ProcessingError` | `Exception` | 处理流程可恢复错误 |
| `IOFailure` | `RuntimeError` | Excel 文件读写失败 |

### 10.4 安全工具

```python
from src.processors.base_processor import sanitize_cell_value

sanitize_cell_value("=cmd|")   # → "'=cmd|"   (Excel 公式注入防护)
sanitize_cell_value("normal")  # → "normal"    (正常内容原样返回)
```

---

## 11. 故障排除

### 配置文件不存在

```
❌ 配置错误: 配置文件不存在: config/categories.yaml
```

确保在项目根目录下运行命令，或用 `--config` 指定路径。

### 未知产品类型

```
❌ 配置错误: 未知产品类型: xyz。可用类型: azd, azh, hd, ...
```

使用 `python -m src.main list-types` 查看可用类型。

### 源文件 / 模板不存在

```
❌ 处理错误: 源文件不存在: data/xxx.xlsx
```

检查文件路径。用 `--source` / `--template` 参数指定正确路径。

### YAML 库未安装

```
❌ 配置错误: 读取 YAML 配置文件需要 PyYAML 库。
```

解决方式二选一：
- `pip install pyyaml`
- 将配置文件转为 JSON 格式（`.json` 扩展名）

### Windows 终端中文/Emoji 乱码

```bash
chcp 65001                       # 切换终端为 UTF-8
set PYTHONIOENCODING=utf-8
```

或使用 Windows Terminal（推荐）。

### Excel 文件被锁定

```
❌ 文件错误: 加载模板失败 [...]: [Errno 13] Permission denied
```

关闭 Excel 中打开的对应文件后重试。

### pandas / openpyxl 版本冲突

```bash
pip install -r requirements.txt  # 使用项目锁定版本
```

---

> **相关文档**：[CLAUDE.md](../CLAUDE.md)（开发者指引）、[docs/CHANGELOG.md](CHANGELOG.md)（变更历史）
