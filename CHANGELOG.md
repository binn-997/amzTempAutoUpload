# CHANGELOG

## [Unreleased] — 2026-06-12

### 重构 (BREAKING)
- 提取 `BaseProcessor` 基类，消除 8 个品牌文件 ~1,672 行重复代码
- 品牌处理器移至 `src/processors/`，每个子类仅 ~10-20 行配置
- 新增 `src/main.py` 统一 CLI 入口（`--brand`, `--input`, `--template`, `--chunk`, `--dry-run`）
- 原 `prov_uploadEcel_Split_*.py` 改为 2 行委托 stub，保持向后兼容

### 新增
- `requirements.txt`：锁定 `pandas==3.0.3`, `openpyxl==3.1.5`, `pyinstaller==6.20.0`
- `CLAUDE.md`：项目架构指引文档
- `REFACTOR_PLAN.md`：重构方案与差异矩阵
- 异常处理：`ProcessingError` / `IOFailure` 异常类，所有 Excel I/O 包裹 try/except
- 安全加固：`_sanitize_cell_value()` 检测公式注入（`= + - @` 前缀自动转义）
- 完整类型注解：`mypy --strict` 零错误

### 优化
- 模板文件只加载一次到内存 BytesIO，分块时复用，避免重复磁盘 I/O
- 尺码替换规则配置化（`SIZE_REPLACEMENTS`）

### 移除
- 删除 `prov_uploadEcel_Split_hd.py` 和 `hh.py` 中注释掉的旧 SRC 配置代码

### 工程化
- 初始化 Git 仓库（`.gitignore` 排除数据文件、构建产物、IDE 配置）
- `.gitignore` 覆盖 `*.xlsx`, `dist/`, `build/`, `.mypy_cache/`, `.venv/` 等
