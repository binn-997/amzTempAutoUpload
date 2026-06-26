# autoUpload — Amazon 商品数据处理工具

> 配置驱动型 Excel 处理工具，用于 Amazon 德国站 Listing 批量上传。

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 列出可用产品类型
python -m src.main list-types

# 处理指定产品类型（使用配置文件默认值）
python -m src.main process --type azd

# 或使用兼容入口（传统方式）
python prov_uploadEcel_Split_azd.py
```

## 项目结构

```
├── src/                           # 核心代码
│   ├── main.py                    # CLI 入口
│   ├── config_loader.py           # 配置加载
│   └── processors/
│       └── base_processor.py      # 通用处理器
├── config/
│   └── categories.yaml            # ★ 产品类型配置（运营人员编辑此文件）
├── temp/                          # Excel 模板文件
├── prov_uploadEcel_Split_*.py     # 兼容入口（8 个产品类型）
├── process.py                     # 颜色翻译工具
├── split_excel.py                 # Excel 拆分工具
└── requirements.txt               # Python 依赖
```

## 详细文档

完整使用说明、配置字段解释、添加新产品类型等请参见：

- **[docs/README_usage.md](docs/README_usage.md)** — ★ 完整使用手册（推荐首先阅读）
- **[docs/CHANGELOG.md](docs/CHANGELOG.md)** — 版本变更日志
- **[CLAUDE.md](CLAUDE.md)** — 开发者指引

## 开发

```bash
pip install -r requirements-dev.txt   # 安装开发依赖（mypy, pytest, ruff）
```
