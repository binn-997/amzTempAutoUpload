# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Amazon 德国站商品数据批量处理工具集。将商品 Excel 源数据经过预处理（颜色英→德翻译、尺码标准化、父子 SKU 继承）后，按 Amazon 模板格式拆分为可上传的分片文件。

## 常用命令

```bash
# 运行颜色翻译工具
python process.py

# 运行单个品牌处理器（修改 Config.SOURCE_FILE 后）
python prov_uploadEcel_Split_azd.py

# 拆分工具 — 拖放/CLI 模式
python split_excel.py                          # 自动检测当前目录 Excel
python split_excel.py -s 源文件.xlsx -n 500     # 指定文件和分块大小

# 打包为 exe（需 PyInstaller）
pyinstaller 拆分表工具.spec
# 或通过 bat 脚本（仅打包 kle 处理器，按需取消注释其他行）
# build_all.bat

# 安装依赖
pip install -r requirements.txt
```

## 架构要点

### 两类脚本

| 类别 | 文件 | 作用 |
|------|------|------|
| **独立工具** | `process.py`, `split_excel.py` | 颜色翻译 & 通用 Excel 拆分 |
| **品牌处理器** | `prov_uploadEcel_Split_*.py` (×8) | 各品牌数据→Amazon 模板映射 |

### 品牌处理器共享结构（95% 代码重复）

8 个 `prov_uploadEcel_Split_*.py` 文件共享完全相同的数据处理流水线，仅 `Config` 类参数不同：

```
Config 类 → preprocess_source_data() → fill_data_to_template()
           → clear_parent_rows() → save_split_workbooks() → main()
```

**每个品牌的差异仅在 Config 类中**：
- `SOURCE_FILE` — 源 Excel 文件名
- `TEMPLATE_FILE` — 对应的 `./temp/prov_*Temp.xlsm` 模板
- `CHUNK_SIZE` — 每文件行数上限（500 或 1000）
- `SIMPLE_MAP` / `MULTI_MAP` — 列映射关系（各品牌模板列位置不同）
- `PARENT_CLEAR_LETTERS` — 父体行需清空的列
- `SIZE_REPLACEMENTS` — 尺码正则替换规则

### 处理流水线

1. `pd.read_excel` 读取源数据（无表头，全字符串）
2. **预处理**：列补齐至 60 列 → NaN 填充 → 列重命名 → 尺码正则替换 → 子体从父体继承 keywords/bullets
3. **模板填充**：按 Config 映射逐行写入模板（1:1、1:N、Bullets、图片 URL）
4. **父体清除**：父体行中清空子体专属字段（价格、尺码等），取消合并单元格
5. **分块保存**：按 CHUNK_SIZE 拆分，每个分块独立加载模板→填充→清除→保存

### 关键外部文件

- `temp/` — Excel 模板文件夹（`.xlsm`/`.xlsx`），运行时依赖，打包时需 `--add-data`
- `prov已上传/` — 已上传文件归档，程序不读取
- `md/mailPrompt.md` — 亚马逊客服 AI Prompt 模板，与 Python 代码无关

## 注意

- 无测试框架、无 linter 配置
- 无类型注解
- Config 中 `SOURCE_FILE` 硬编码，运行前需手动修改
- 输出目录 `prov_output/` 中的文件会被直接覆盖，无备份
- 所有 Excel I/O 无异常处理
