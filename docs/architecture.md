# 架构与开发

`alt` 是唯一的命令入口。CLI 负责参数与错误码；配置加载器负责 YAML/JSON 解析、合并和校验；领域模块负责转换规则；Excel 层只负责读取、写入和模板复制。

产品类型不再通过 Python 子类或脚本复制扩展。新增类型时，在 `config/categories.yaml` 中添加条目，并为真实模板执行人工验收。

## 配置

`defaults` 保存共享规则，`categories` 覆盖差异。字典递归合并，列表整体替换。配置加载阶段会校验顶级结构、品类条目和模板列字母；业务模板兼容性需要在真实资产上确认。

## 开发

安装开发依赖后运行：

```bash
ruff check .
mypy src/amazon_listing_toolkit
pytest
```

测试使用临时生成的工作簿，绝不依赖或提交真实商品资料。
