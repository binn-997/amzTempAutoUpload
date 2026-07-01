# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Amazon 德国站商品数据批量处理工具集。配置驱动型架构：YAML 定义产品类型 → `config_loader` 解析合并 → `BaseProcessor` 执行流水线。支持 8 种服装产品类型（西装、裤装、连衣裙、连裤袜、上装），按 Amazon 库存文件模板格式输出可分批上传的 Excel。

## 常用命令

```bash
# 安装依赖
pip install -r requirements.txt

# === 核心 CLI（推荐） ===
python -m src.main list-types                     # 列出所有可用产品类型
python -m src.main process --type azd             # 处理指定产品类型（使用配置默认值）
python -m src.main process --type th --dry-run    # 仅预览，不写文件
python -m src.main process --type hd --source data/other.xlsx --chunk 800

# === 颜色翻译 ===
python -m src.main translate --source data/701th.xlsx
python -m src.main translate --source data/xxx.xlsx --no-translate   # 仅变体检测
python -m src.main translate --source data/xxx.xlsx --color-col E --size-col F

# === 利润定价 ===
python -m src.main price --source data/630td.xlsx
python -m src.main price --source data/630td.xlsx --target 0.25 --rate 7.8
python -m src.main price --source data/630td.xlsx --station US --rate 7.0

# === 模块自测 ===
python src/color_translator.py      # 颜色翻译 + 变体检测单元测试
python src/profit_calculator.py     # 利润计算器验证测试

# === 兼容入口（等同于 python -m src.main process --type <xx>） ===
python prov_uploadEcel_Split_azd.py
```

## 架构要点

### 配置驱动模型（v2.0+）

整个系统由 [config/categories.yaml](config/categories.yaml) 驱动：

```
categories.yaml          config_loader.py           BaseProcessor
┌──────────────┐        ┌──────────────────┐       ┌────────────────┐
│ defaults     │   →    │ resolve_category  │  →   │ __init__()     │
│ source_cols  │        │ _deep_merge()     │      │ extracts all   │
│ template_cols│        │ inject shared cols│      │ params from cfg│
│ categories:  │        └──────────────────┘       └────────────────┘
│   azd: ...   │                                            │
│   th:  ...   │        8 个 prov_uploadEcel_Split_*.py     │
└──────────────┘        是薄兼容入口，内部委托 CLI          │
                                                           ▼
                                              run() → _load_and_preprocess()
                                                   → save_split_workbooks()
```

### 处理流水线（BaseProcessor）

1. `_load_and_preprocess()` — 读取源 Excel（无表头，全字符串）→ 列补齐 60 列 → NaN 填充 → 列重命名 → 尺码正则替换 → 子体从父体继承 keywords/bullets → `post_preprocess()` 钩子
2. `fill_data_to_template()` — 按配置映射逐行写入：1:1 列（`simple_map`）→ 1:N 列（`multi_map`）→ Bullet Points → 图片 URL → 父子标识
3. `clear_parent_rows()` — 父体行清空子体专属字段（价格、尺码等），取消合并单元格
4. `save_split_workbooks()` — 按 `chunk_size` 拆分，模板只加载一次到内存 `BytesIO`，后续分块从缓存创建

### 配置合并规则（`_deep_merge`）

- **标量值**：产品类型覆盖 defaults
- **字典**（如 `size_replacements`）：递归合并，产品类型键优先
- **列表**（如 `parent_clear_letters`）：产品类型完全替换 defaults

### 新增模块（v2.1）

| 模块 | 作用 |
|------|------|
| [src/color_translator.py](src/color_translator.py) | 颜色英→德翻译 + 按 parent_sku 分组检测同款变体，签名 `(color, sizes_tuple)` 重复时加 V1/V2 后缀 |
| [src/profit_calculator.py](src/profit_calculator.py) | 多站点（US/CA/DE）利润核算，按售价区间分档扣点率，重量→运费查表，支持反算最低售价 |

### 关键外部依赖

- `data/` — 源 Excel 文件（`.xlsx`），gitignore 排除
- `temp/` — Excel 模板文件（`.xlsm`/`.xlsx`），运行时依赖，打包需 `--add-data`
- `prov_output/` — 输出目录，文件会被直接覆盖，无备份
- `config/categories.yaml` — 单文件配置中心，运营人员编辑此文件即可添加新产品类型

## 注意

- 无 pytest/unittest 测试框架；自测内嵌在 `color_translator.py` 和 `profit_calculator.py` 的 `if __name__ == "__main__"` 块中
- 所有 Excel I/O 无异常处理重试逻辑
- 8 个 `prov_uploadEcel_Split_*.py` 是兼容入口（thin wrappers），所有逻辑在 `BaseProcessor` 中，修改流水线只需改基类
