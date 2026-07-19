# 开发者指引

项目通过 `pyproject.toml` 管理依赖和工具。安装后使用 `alt --help` 查看命令；不要恢复旧品牌脚本或 `src.main` 入口。

```bash
python -m pip install -e ".[dev]"
ruff check .
mypy src/amazon_listing_toolkit
pytest
```

真实 Excel、模板和输出必须保持在 Git 之外。产品差异应写入包内的 `config/categories.yaml`，而不是复制处理器代码。
