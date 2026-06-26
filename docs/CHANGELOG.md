# CHANGELOG

## [v2.0] — 2026-06-26 — 第二次重构：配置驱动化

### 重构目标
1. **完全命令行驱动**：不修改任何 `.py` 文件即可切换源文件、模板、分块大小等参数
2. **映射规则配置化**：产品类型差异从 Python 代码移至 YAML/JSON 配置文件
3. **清理处理器代码**：删除 8 个子类文件，改为通用处理器 + 配置文件
4. **保持向后兼容**：`prov_uploadEcel_Split_*.py` 兼容入口行为不变

### 新增
- `config/categories.yaml`：产品类型配置文件（YAML 格式，支持 JSON）
  - 包含 8 个产品类型（azd/azh/hd/hh/kle/std/td/th）的完整参数
  - 全局 `defaults` 节减少重复，`categories` 节定义各类型差异化配置
  - 支持 `size_replacements` 字典合并（产品类型只需定义额外规则）
  - 包含 `source_columns` 和 `template_columns` 共享映射配置
- `src/config_loader.py`：配置加载 & 合并工具模块
  - `load_config(path)`：根据扩展名自动选择 YAML/JSON 解析器
  - `resolve_category_config(config, key)`：深度合并 defaults + 产品类型配置
  - `list_categories(config)`：列出所有可用产品类型
  - `find_project_root()`：自动定位项目根目录
  - `ConfigError`：配置相关异常类
- `README_usage.md`：完整使用文档
  - 安装依赖说明
  - 配置文件结构 & 字段说明
  - 命令行示例（含常用场景）
  - 添加新产品类型步骤
  - FAQ

### 修改
- `src/processors/base_processor.py`：重构为配置驱动
  - `__init__(self, config: dict)` 接受配置字典，所有参数从 `self._config` 注入
  - 原类属性保留为回退默认值
  - `simple_map`、`multi_map` 等计算属性基于注入参数动态计算
  - 新增钩子方法 `post_preprocess()` 和 `post_fill_row()`（默认空操作）
  - `_sanitize_cell_value` 改为模块级函数 `sanitize_cell_value`，供外部工具复用
  - Windows 终端编码修复（`sys.stdout.reconfigure(encoding='utf-8')`）
- `src/main.py`：重写为子命令式 CLI
  - `process` 子命令：执行数据处理流水线
    - `--type`（必选）：产品类型标识
    - `--source`（可选）：源文件，覆盖配置默认值
    - `--template`（可选）：模板文件，覆盖配置默认值
    - `--chunk`（可选）：分块行数，覆盖配置默认值
    - `--output-dir`（可选）：输出目录，覆盖配置默认值
    - `--config`（可选）：配置文件路径（默认 `config/categories.yaml`）
    - `--dry-run`（可选）：仅预览不写文件
  - `list-types` 子命令：列出所有可用产品类型
  - 移除旧版 `BRAND_REGISTRY` 硬编码，改为配置文件动态加载
  - 添加 `sys.path` 修复和编码修复
- `prov_uploadEcel_Split_*.py` (×8)：重写兼容入口
  - 从配置文件读取对应产品类型默认值
  - 委托 `BaseProcessor` 执行
  - 每个文件仅需修改 `_TYPE_KEY` 变量即可适配
- `src/__init__.py`：更新导出（`BaseProcessor`, `ProcessingError`, `IOFailure`）
- `src/processors/__init__.py`：更新模块说明
- `requirements.txt`：添加 `pyyaml==6.0.3`（可选，JSON 配置文件无需安装）

### 删除
- `src/processors/azd.py` ~ `th.py` (×8)：子类文件已由配置文件完全替代
  - 删除 ~30 行重复类定义代码（8 文件 × 约 4 行/文件 ≈ 32 行，不含注释）
- 相关的 `__pycache__/` 字节码缓存

### 验证结果
- ✅ 所有 Python 文件语法检查通过
- ✅ YAML 配置加载 & 合并正确（8 个产品类型均验证通过）
- ✅ th 的 12 条尺码规则合并正确（8 默认 + 4 特有）
- ✅ hh 的 `include_final_keywords=False` 正确生效
- ✅ std 的 `size_class_col=None` 正确生效
- ✅ CLI `list-types` 子命令正常输出
- ✅ CLI `process --type azd --dry-run` 正常执行
- ✅ 兼容入口 `python prov_uploadEcel_Split_azd.py` 正常运行

### 架构变化
```
重构前:
  src/processors/
    base_processor.py (481行, 含类属性硬编码)
    azd.py ~ th.py (8个, 各3~28行, 硬编码 SOURCE_FILE/TEMPLATE_FILE)
  src/main.py (BRAND_REGISTRY 硬编码8个类的映射)
  prov_uploadEcel_Split_*.py (导入子类 → 调用 .run())

重构后:
  config/
    categories.yaml (★ 单一真相源)
  src/
    config_loader.py (配置加载/合并)
    processors/
      base_processor.py (★ 通用处理器, 配置注入)
      __init__.py (仅导出, 无子类)
    main.py (配置驱动的 CLI)
  prov_uploadEcel_Split_*.py (配置→BaseProcessor)
```

### 向后兼容性
- `python prov_uploadEcel_Split_azd.py` 等 8 个入口行为不变
- 输出目录 `prov_output/` 和文件名格式不变
- 模板文件 `temp/` 结构和内容不变
- `process.py`、`split_excel.py` 未接触

---

## [v1.0] — 2026-06-12 — 第一次重构：提取基类

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
