# 利润计算器 — 使用手册

> **模块**: `src/profit_calculator.py` | **版本**: v1.0 | **更新日期**: 2026-06-30

---

## 目录

1. [概述](#1-概述)
2. [快速开始](#2-快速开始)
3. [核心公式](#3-核心公式)
4. [API 参考](#4-api-参考)
5. [使用场景](#5-使用场景)
6. [运费查找表](#6-运费查找表)
7. [自定义扣点率](#7-自定义扣点率)
8. [批量计算](#8-批量计算)
9. [在其它脚本中调用](#9-在其它脚本中调用)

---

## 1. 概述

`ProfitCalculator` 是一个多站点（US / CA / DE）跨境电商利润计算模块，实现与 `data/lr.xlsx` 利润计算表完全一致的计算逻辑。

### 支持的站点

| 站点 | 枚举值 | 模式 | 扣点率档位 |
|------|--------|------|-----------|
| 德国 | `Station.DE` | FBM 自发货 | 2 档（15 EUR 分界） |
| 美国 | `Station.US` | FBA | 3 档（15/20 USD 分界） |
| 加拿大 | `Station.CA` | FBA | 2 档（20 CAD 分界） |

### 核心功能

| 功能 | 方法 | 说明 |
|------|------|------|
| 单次利润计算 | `calculate()` | 输入售价/成本/运费，返回毛利率和净利润 |
| 重量→运费自动查找 | `calculate_from_weight()` | 只需提供包裹重量，自动查运费表 |
| 批量计算 | `calculate_batch()` | 传入 DataFrame，批量追加利润列 |
| 投入回报率 | `roi()` | ROI = 净利润 / 总成本 |
| 反算最低售价 | `breakeven_price()` | 给定目标毛利率，反求最低售价 |
| 可读摘要 | `summary()` | 单次计算结果的完整中文/英文摘要 |

---

## 2. 快速开始

```python
from src.profit_calculator import ProfitCalculator, Station, quick_de

# ===== 方式 1: 完整参数（最灵活） =====
calc = ProfitCalculator(station=Station.DE)
result = calc.calculate(
    item_price=49.99,         # 商品售价（外币）
    shipping_price=6.99,      # 向买家收取的运费（外币）
    exchange_rate=7.8,        # 汇率（RMB / 外币）
    sku_cost_rmb=130.0,       # SKU 成本（人民币）
    shipping_cost_rmb=43.5,   # 运费成本（人民币）
)
print(f"毛利率: {result.margin_rate:.2%}")    # → 16.46%
print(f"净利润: RMB {result.profit_rmb:.2f}") # → RMB 73.17

# ===== 方式 2: 通过重量自动查运费 =====
result = calc.calculate_from_weight(
    item_price=49.99,
    weight_g=160,             # 包裹重量 160g，自动查运费表
    exchange_rate=7.8,
    sku_cost_rmb=130.0,
)
print(f"运费: RMB {result.shipping_cost_rmb:.2f}")  # → RMB 29.00

# ===== 方式 3: 一键快捷函数 =====
result = quick_de(
    item_price=49.99,
    sku_cost_rmb=130.0,
    weight_g=160,
)
```

---

## 3. 核心公式

### 统一公式

```
毛利率 = (总售价×汇率 - SKU成本 - 运费成本 - 总售价×汇率×综合扣点率) / (总售价×汇率)

即:
margin = (total_rmb - sku_cost - shipping_cost - total_rmb × deduction_rate) / total_rmb
```

### 各站点综合扣点率拆解

#### 德国站 DE（FBM 自发货）

| 售价区间 | 综合扣点率 | VAT | 佣金 | 广告 | 退货损耗 |
|----------|:----------:|:---:|:----:|:----:|:--------:|
| 0-15 EUR | **37.5%** | 16% | 8% | 10% | 3.5% |
| 15 EUR+ | **44.5%** | 16% | 15% | 10% | 3.5% |

> **设计意图**：FBM 模式下，15 EUR 以上佣金从 8% 跳到 15%（+7%），扣点率相应从 37.5% 升至 44.5%。雷打不动的 13.5%（广告 10% + 退货 3.5%）在选品阶段就提前扣掉，防止毛利率虚高。

#### 美国站 US（FBA）

| 售价区间 | 综合扣点率 | 佣金 | 广告 | 退货损耗 |
|----------|:----------:|:----:|:----:|:--------:|
| 0-15 USD | **28%** | 17% | 8% | 3% |
| 15-20 USD | **33%** | 17% | 12% | 4% |
| 20 USD+ | **40%** | 17% | 17% | 6% |

#### 加拿大站 CA（FBA）

| 售价区间 | 综合扣点率 | 佣金 | 广告 | 退货损耗 |
|----------|:----------:|:----:|:----:|:--------:|
| 0-20 CAD | **40%** | 15% | 17% | 8% |
| 20 CAD+ | **47%** | 15% | 22% | 10% |

### 与 Excel 公式对照验证

以下数据来自 `data/lr.xlsx` 实测：

| 站点 | 输入数据 | Excel 公式 | 代码输出 | 匹配 |
|------|----------|-----------|----------|:--:|
| DE (0-15) | 8.99+5.96, 汇率7.8, 成本10+26 | `0.316279` | `0.316279` | ✅ |
| DE (15+) | 49.99+6.99, 汇率7.8, 成本130+43.5 | `0.164625` | `0.164625` | ✅ |
| US | 12.99+4.99, 汇率7, 成本25+42.5 | `0.183690` | `0.183690` | ✅ |
| CA | 9.99+0, 汇率5, 成本8+23.5 | `-0.030631` | `-0.030631` | ✅ |

---

## 4. API 参考

### 4.1 `ProfitCalculator`

```python
class ProfitCalculator:
    def __init__(
        self,
        station: Station = Station.DE,
        shipping_table: ShippingTable | None = None,
        use_detail: bool = True,
    ):
        ...
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `station` | `Station` | `Station.DE` | 目标站点（US / CA / DE） |
| `shipping_table` | `ShippingTable \| None` | `None` | 运费查找表，None 则使用内置表 |
| `use_detail` | `bool` | `True` | 是否在结果中附加扣点明细（VAT/佣金/广告/退货） |

### 4.2 `calculate()`

```python
def calculate(
    self,
    item_price: float,
    shipping_price: float = 0.0,
    exchange_rate: float = 7.8,
    sku_cost_rmb: float = 0.0,
    shipping_cost_rmb: float = 0.0,
    *,
    weight_g: float | None = None,
    shipping_method: str = "ceil",
) -> ProfitResult:
    ...
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|:--:|------|
| `item_price` | float | ✅ | 商品售价（外币，不含运费） |
| `shipping_price` | float | ❌ | 向买家收取的运费（外币） |
| `exchange_rate` | float | ❌ | 汇率（RMB / 外币），默认 7.8 |
| `sku_cost_rmb` | float | ❌ | SKU 成本（人民币） |
| `shipping_cost_rmb` | float | ❌ | 运费成本（人民币）。若提供 `weight_g` 则忽略 |
| `weight_g` | float | ❌ | 包裹重量（克），提供则自动查运费表 |
| `shipping_method` | str | ❌ | 查表方法：`"linear"`（插值）\| `"ceil"`（向上）\| `"floor"`（向下） |

**返回值**: `ProfitResult`

### 4.3 `ProfitResult`

```python
@dataclass
class ProfitResult:
    # ── 输入 ──
    station: Station           # 站点
    item_price: float          # 商品售价（外币）
    shipping_price: float      # 运费收入（外币）
    exchange_rate: float       # 汇率
    sku_cost_rmb: float        # SKU 成本（RMB）
    shipping_cost_rmb: float   # 运费成本（RMB）

    # ── 中间量 ──
    total_price_foreign: float # 总售价（外币）
    total_price_rmb: float     # 总售价（RMB）
    deduction_rate: float      # 综合扣点率
    deduction_tier_label: str  # 命中档位名（如 "15 EUR+"）
    deduction_amount_rmb: float# 综合扣除金额（RMB）
    cost_rmb_total: float      # 总成本（RMB）

    # ── 结果 ──
    profit_rmb: float          # 净利润（RMB）
    margin_rate: float         # 毛利率

    # ── 明细 ──
    detail: dict | None        # 扣点明细 {VAT: xx, Commission: xx, ...}
```

### 4.4 其它方法

```python
# 投入回报率
def roi(self, **kwargs) -> float:
    """ROI = 净利润 / 总成本。kwargs 同 calculate()"""

# 反算最低售价
def breakeven_price(
    self,
    exchange_rate: float = 7.8,
    sku_cost_rmb: float = 0.0,
    shipping_cost_rmb: float = 0.0,
    shipping_price: float = 0.0,
    target_margin: float = 0.15,
    *,
    weight_g: float | None = None,
) -> float:
    """
    给定目标毛利率，反求最低售价（外币）。
    使用迭代法自动匹配正确的扣点率档位。
    """

# 可读摘要
def summary(self, result: ProfitResult) -> str:
    """生成单次计算结果的格式化文本摘要。"""

# 批量计算
def calculate_batch(
    self,
    df: pd.DataFrame,
    *,
    item_price_col: str = "item_price",
    ...
) -> pd.DataFrame:
    """批量计算，在原 DataFrame 基础上追加利润列。"""
```

---

## 5. 使用场景

### 场景 1：日常核价

上线新产品前，判断利润是否达标：

```python
from src.profit_calculator import ProfitCalculator, Station

calc = ProfitCalculator(Station.DE)

# 新品参数
r = calc.calculate_from_weight(
    item_price=39.99,
    weight_g=200,
    sku_cost_rmb=85.0,
    exchange_rate=7.8,
)

print(calc.summary(r))
# 输出:
# ========================================
#   DE Profit Analysis
# ========================================
# Station:        DE
# Item Price:     39.99
# ...
# Net Profit:     xx.xx RMB
# Margin Rate:    xx.xx%
```

### 场景 2：反算定价

老板要求毛利率不低于 25%，这款产品成本 RMB 85，运费 RMB 30.5，该定多少价？

```python
min_price = calc.breakeven_price(
    sku_cost_rmb=85.0,
    shipping_cost_rmb=30.5,
    shipping_price=6.99,  # 收买家运费
    target_margin=0.25,
)
print(f"最低售价: EUR {min_price:.2f}")
```

### 场景 3：多站点比价

同一款产品，对比在 US 和 DE 哪个站更赚钱：

```python
for station in [Station.US, Station.DE]:
    calc = ProfitCalculator(station)
    r = calc.calculate_from_weight(
        item_price=29.99, weight_g=150,
        exchange_rate=7.0 if station == Station.US else 7.8,
        sku_cost_rmb=65.0,
    )
    print(f"{station.value}: margin={r.margin_rate:.2%}, "
          f"profit=RMB {r.profit_rmb:.2f}")
```

### 场景 4：ROI 评估

判断投入产出是否合理：

```python
calc = ProfitCalculator(Station.DE)
roi = calc.roi(
    item_price=29.99, weight_g=120,
    sku_cost_rmb=55.0, exchange_rate=7.8,
)
print(f"ROI: {roi:.2%}")
# ROI = 净利润 / (SKU成本 + 运费) → 越高越好
```

---

## 6. 运费查找表

### 内置运费表

每个站点预置了完整的重量→运费映射表（来自 `data/lr.xlsx`）。

| 站点 | 重量范围 | 运费范围（RMB） |
|------|----------|----------------|
| DE | 80-400g | 26.0-41.5 |
| US | 80-400g | 31.0-52.5 |
| CA | 80-400g | 21.0-39.5 |

### 查表方法

| 方法 | 说明 | 示例（重量 150g） |
|------|------|-------------------|
| `"ceil"` | 取 ≥ 输入重量的最小档位 | 150g → 查 160g 档 |
| `"floor"` | 取 ≤ 输入重量的最大档位 | 150g → 查 140g 档 |
| `"linear"` | 区间线性插值 | 150g → 相邻档位插值 |

```python
# 默认使用 ceil（更保守，运费宁高勿低）
r = calc.calculate_from_weight(
    item_price=29.99, weight_g=150,
    shipping_method="ceil",
)

# 使用线性插值（更精确）
r = calc.calculate_from_weight(
    item_price=29.99, weight_g=150,
    shipping_method="linear",
)
```

### 从 Excel 文件加载自定义运费表

```python
from src.profit_calculator import ShippingTable, Station

# 从利润计算 Excel 加载
table = ShippingTable.from_excel(
    "data/lr.xlsx",
    station=Station.DE,
    weight_row=9,   # 重量（g）在第 9 行
    cost_row=10,    # 运费（元）在第 10 行
)

calc = ProfitCalculator(Station.DE, shipping_table=table)
```

### 手动定义运费表

```python
from src.profit_calculator import ShippingTable, WeightShippingRow

custom_table = ShippingTable([
    WeightShippingRow(100, 25.0),
    WeightShippingRow(200, 30.0),
    WeightShippingRow(300, 35.0),
    WeightShippingRow(400, 40.0),
    WeightShippingRow(500, 45.0),
])

calc = ProfitCalculator(Station.DE, shipping_table=custom_table)
```

---

## 7. 自定义扣点率

如果实际的广告/退货费率与内置默认值不同，可以覆盖扣点率：

### 临时替换所有档位

```python
from src.profit_calculator import DeductionTier

calc = ProfitCalculator(Station.DE)

# 自定义一套新档位（完全替换内置配置）
calc.set_custom_tiers([
    DeductionTier(max_price=10.0, rate=0.35, label="<10 EUR"),
    DeductionTier(max_price=20.0, rate=0.40, label="10-20 EUR"),
    DeductionTier(max_price=None, rate=0.48, label="20+ EUR"),
])

# 之后所有计算都使用自定义扣点率
r = calc.calculate(item_price=15.99, ...)
```

### A/B 测试不同广告预算

```python
# 对比 10% vs 15% 广告预算下的利润
from copy import deepcopy

base_tiers = calc.tiers

# 方案 A: 正常广告 10%
calc.set_custom_tiers(base_tiers)
r_a = calc.calculate_from_weight(item_price=29.99, weight_g=150, sku_cost_rmb=55)

# 方案 B: 激进广告 15%（每个档位 +5%）
aggressive_tiers = [
    DeductionTier(t.max_price, t.rate + 0.05, t.label + " (aggressive)")
    for t in base_tiers
]
calc.set_custom_tiers(aggressive_tiers)
r_b = calc.calculate_from_weight(item_price=29.99, weight_g=150, sku_cost_rmb=55)

print(f"正常广告: margin={r_a.margin_rate:.2%}")
print(f"激进广告: margin={r_b.margin_rate:.2%}")
print(f"差异:     {(r_a.margin_rate - r_b.margin_rate):.2%}")
```

---

## 8. 批量计算

对整批 SKU 进行批量利润核算：

```python
import pandas as pd
from src.profit_calculator import ProfitCalculator, Station

# 准备数据（每行一个 SKU）
df = pd.DataFrame({
    "item_price":    [49.99, 29.99, 19.99],
    "shipping_price":[6.99,  5.99,  4.99],
    "exchange_rate": [7.8,   7.8,   7.8],
    "sku_cost_rmb":  [130.0, 65.0,  45.0],
    "weight_g":      [160,   120,   100],
})

calc = ProfitCalculator(Station.DE)

# 批量计算
result_df = calc.calculate_batch(df)

# 结果列:
#   total_price_foreign, total_price_rmb,
#   deduction_rate, deduction_tier,
#   profit_rmb, margin_rate,
#   cost_rmb_total, deduction_amount_rmb

print(result_df[["item_price", "margin_rate", "profit_rmb"]])
#    item_price  margin_rate  profit_rmb
# 0       49.99       0.1646       73.17
# 1       29.99       0.xxxx       xx.xx
# 2       19.99       0.xxxx       xx.xx
```

### 自定义列名映射

如果 DataFrame 列名与默认值不同，可通过参数指定：

```python
result_df = calc.calculate_batch(
    df,
    item_price_col="售价欧元",
    sku_cost_col="采购成本",
    weight_col="重量克",
    # shipping_price_col, exchange_rate_col 等同理
)
```

---

## 9. 在其它脚本中调用

### 导入方式

```python
# 方式 A: 从 src 包导入（推荐）
from src.profit_calculator import ProfitCalculator, Station, quick_de

# 方式 B: 直接导入模块
from src import profit_calculator as pc
result = pc.quick_de(item_price=29.99, sku_cost_rmb=55.0, weight_g=150)
```

### 在处理流水线中集成

```python
# 在品牌处理器中，预处理后追加利润列
import pandas as pd
from src.profit_calculator import ProfitCalculator, Station

class MyProcessor(BaseProcessor):
    def post_preprocess(self, df: pd.DataFrame) -> pd.DataFrame:
        """预处理后自动计算每条 SKU 的预估利润。"""
        calc = ProfitCalculator(Station.DE)

        # 假设 df 中有 price、weight 列
        result_df = calc.calculate_batch(
            df,
            item_price_col="standprice",
            weight_col="package_weight",
            sku_cost_col="sku_cost_rmb",  # 需要事先有成本列
            exchange_rate_col="exchange_rate",
        )
        return result_df
```

### 命令行快速验证

```bash
# 运行自测（验证所有公式与 Excel 匹配）
python -m src.profit_calculator
```

---

## 附录：扣点率设计原理

### 为什么毛利率要「提前扣掉」广告和退货？

跨境电商运营中，如果核价时只计算 `售价 - 成本 - 佣金 - 税 = 利润`，算出的毛利率会**虚高**，因为忽略了两个隐性成本：

1. **广告成本**（ACOS / TACOS）：不开广告没有流量，ACOS 10-15% 是常态。
2. **退货/妥投失败损耗**：跨境退货成本极高（退回国内运费 > 货值本身，通常只能弃置），损失必须平摊。

### DE 站公式的设计哲学

37.5% 和 44.5% 是**风险前置型**核价公式：

- 如果扣掉必扣项（VAT 16% + 佣金 8%/15%）后还有 13.5%，这是故意留给广告和退货的「缓冲区」。
- 如果一个单品在这种「猛药」公式下还能算出 20%+ 的毛利率，它就是真正的**暴利好款**，可以放心投入。
- 反之，如果只能勉强保本，说明这个产品盈利能力不足，需要重新评估定价或成本。

---

> **相关文档**: [README_usage.md](README_usage.md)（项目整体使用手册）、[CLAUDE.md](../CLAUDE.md)（开发者指引）
