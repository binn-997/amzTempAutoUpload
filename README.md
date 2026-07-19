# Amazon Listing Toolkit

> 配置驱动的 Amazon Listing Excel 本地处理工具：模板填充、颜色翻译、利润定价与工作簿拆分。

本项目是个人技术作品集，**不是 Amazon 官方产品**。仓库不包含商品数据、上传模板、宏文件或输出文件；请始终使用副本验证结果后再上传。

## 特性

- YAML 声明式产品类型：列映射、尺码规则和父体字段清理无需新增 Python 类。
- 单一 CLI `alt`：处理、翻译、定价、拆分共用一致的参数和错误输出。
- 父/子 SKU 继承、德语颜色翻译、变体消歧、售价反算与 Excel 公式注入转义。
- `dataclass` 配置校验、类型检查、Ruff 与 pytest 本地质量门槛。

## 安装

需要 Python 3.11–3.14。

```bash
git clone https://github.com/binn-997/amzTempAutoUpload.git
cd amzTempAutoUpload
python -m pip install -e ".[dev]"
```

## 快速开始

```bash
alt list-types
alt process --type azd --source data/source.xlsx --template temp/template.xlsm
alt process --type th --dry-run
alt translate --source data/source.xlsx
alt price --source data/source.xlsx --station DE --target 0.25
alt split --source data/source.xlsx --chunk 500
```

`translate` 与 `price` 会原地回写源工作簿；`process` 和 `split` 会在指定输出目录创建文件。

## 架构

```text
CLI (alt) → 配置加载与校验 → 领域服务 → Excel I/O
                         ├─ 模板处理
                         ├─ 颜色翻译
                         ├─ 利润定价
                         └─ 工作簿拆分
```

产品差异仅存在于 `src/amazon_listing_toolkit/config/categories.yaml`。详细说明见 [架构与开发说明](docs/architecture.md)。

## 本地质量检查

```bash
ruff check .
mypy src/amazon_listing_toolkit
pytest
```

## 许可证

本项目采用 [MIT License](LICENSE)。
