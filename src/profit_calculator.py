"""
跨境电商利润计算器 — 多站点、模块化、可配置。

支持站点：US（美国）、CA（加拿大）、DE（德国 FBM）。
每个站点的综合扣点率按售价区间分段，公式统一为：
    毛利率 = (总售价 - 成本/汇率 - 运费/汇率 - 总售价 × 综合扣点率) / 总售价

综合扣点率包含：VAT/税 + Amazon 佣金 + 广告成本 + 退货损耗。

用法:
    from src.profit_calculator import ProfitCalculator, Station

    calc = ProfitCalculator(station=Station.DE)
    result = calc.calculate(
        item_price=49.99,       # 商品售价（外币）
        shipping_price=6.99,    # 运费收入（外币）
        exchange_rate=7.8,      # 汇率
        sku_cost_rmb=130.0,     # SKU成本（人民币）
        shipping_cost_rmb=43.5,  # 运费成本（人民币）
    )
    print(f"毛利率: {result.margin_rate:.2%}")
    print(f"净利润: ¥{result.net_profit_rmb:.2f}")
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Sequence

import pandas as pd

logger = logging.getLogger(__name__)


# ==================== 枚举 ====================


class Station(str, Enum):
    """支持的亚马逊站点。"""

    US = "US"   # 美国
    CA = "CA"   # 加拿大
    DE = "DE"   # 德国（FBM 自发货）


# ==================== 数据类 ====================


@dataclass
class DeductionTier:
    """一个扣点率档位。

    Args:
        max_price: 该档位的售价上限（外币，不含运费），None 表示无上限。
        rate: 综合扣点率（0.375 表示 37.5%）。
        label: 档位名称（便于显示和调试）。
    """

    max_price: Optional[float]
    rate: float
    label: str


@dataclass
class ProfitResult:
    """单次利润计算结果。"""

    # ── 输入 ──
    station: Station
    item_price: float           # 商品售价（外币）
    shipping_price: float       # 运费收入（外币）
    exchange_rate: float        # 汇率（本币/外币）
    sku_cost_rmb: float         # SKU 成本（人民币）
    shipping_cost_rmb: float    # 运费成本（人民币）

    # ── 中间量 ──
    total_price_foreign: float = 0.0        # 总售价（外币，含运费）
    total_price_rmb: float = 0.0            # 总售价（人民币）
    deduction_rate: float = 0.0             # 命中的综合扣点率
    deduction_tier_label: str = ""          # 命中的档位名称
    deduction_amount_rmb: float = 0.0       # 综合扣除总额（人民币）
    cost_rmb_total: float = 0.0             # 总成本（SKU + 运费，人民币）

    # ── 结果 ──
    profit_rmb: float = 0.0                 # 净利润（人民币）
    margin_rate: float = 0.0                # 毛利率（0.30 表示 30%）

    # ── 明细拆解（可选） ──
    detail: Optional[Dict[str, float]] = None  # 扣点明细（VAT、佣金、广告、退货）

    def __repr__(self) -> str:
        return (
            f"ProfitResult(station={self.station.value}, "
            f"margin={self.margin_rate:.2%}, "
            f"profit=RMB{self.profit_rmb:.2f}, "
            f"tier={self.deduction_tier_label})"
        )


# ==================== 站点扣点配置 ====================


class StationConfig:
    """定义每个站点的扣点率档位。

    综合扣点率 = VAT/税 + Amazon 佣金 + 广告 + 退货损耗。

    DE (FBM) 拆解:
      - 0-15€ 综合 37.5% = VAT 16% + 佣金 8% + 广告 10% + 退货 3.5%
      - 15€+  综合 44.5% = VAT 16% + 佣金 15% + 广告 10% + 退货 3.5%

    US 拆解:
      - 0-15$   综合 28% = 佣金 17% + 广告 8% + 退货 3%
      - 15-20$  综合 33% = 佣金 17% + 广告 12% + 退货 4%
      - 20$+    综合 40% = 佣金 17% + 广告 17% + 退货 6%

    CA 拆解:
      - 0-20 CAD 综合 40% = 佣金 15% + 广告 17% + 退货 8%
      - 20+ CAD  综合 47% = 佣金 15% + 广告 22% + 退货 10%
    """

    # ── DE: FBM 自发货 ──
    DE_TIERS: List[DeductionTier] = [
        DeductionTier(max_price=15.0, rate=0.375, label="0-15 EUR"),
        DeductionTier(max_price=None, rate=0.445, label="15 EUR+"),
    ]

    DE_DETAIL: Dict[str, Dict[str, float]] = {
        "0-15 EUR": {"VAT": 0.16, "Commission": 0.08, "Ads": 0.10, "Returns": 0.035},
        "15 EUR+": {"VAT": 0.16, "Commission": 0.15, "Ads": 0.10, "Returns": 0.035},
    }

    # ── US: FBA ──
    US_TIERS: List[DeductionTier] = [
        DeductionTier(max_price=15.0, rate=0.28, label="0-15 USD"),
        DeductionTier(max_price=20.0, rate=0.33, label="15-20 USD"),
        DeductionTier(max_price=None, rate=0.40, label="20 USD+"),
    ]

    US_DETAIL: Dict[str, Dict[str, float]] = {
        "0-15 USD": {"Commission": 0.17, "Ads": 0.08, "Returns": 0.03},
        "15-20 USD": {"Commission": 0.17, "Ads": 0.12, "Returns": 0.04},
        "20 USD+": {"Commission": 0.17, "Ads": 0.17, "Returns": 0.06},
    }

    # ── CA: FBA ──
    CA_TIERS: List[DeductionTier] = [
        DeductionTier(max_price=20.0, rate=0.40, label="0-20 CAD"),
        DeductionTier(max_price=None, rate=0.47, label="20 CAD+"),
    ]

    CA_DETAIL: Dict[str, Dict[str, float]] = {
        "0-20 CAD": {"Commission": 0.15, "Ads": 0.17, "Returns": 0.08},
        "20 CAD+": {"Commission": 0.15, "Ads": 0.22, "Returns": 0.10},
    }

    @classmethod
    def get_tiers(cls, station: Station) -> List[DeductionTier]:
        """获取指定站点的扣点率档位列表。"""
        mapping = {
            Station.US: cls.US_TIERS,
            Station.CA: cls.CA_TIERS,
            Station.DE: cls.DE_TIERS,
        }
        return mapping[station]

    @classmethod
    def get_detail(cls, station: Station, tier_label: str) -> Dict[str, float]:
        """获取指定站点/档位的扣点明细。"""
        mapping = {
            Station.US: cls.US_DETAIL,
            Station.CA: cls.CA_DETAIL,
            Station.DE: cls.DE_DETAIL,
        }
        return mapping[station].get(tier_label, {})


# ==================== 运费查找表 ====================


@dataclass
class WeightShippingRow:
    """重量-运费映射表中的一行。

    Args:
        weight_g: 重量（克）
        shipping_cost_rmb: 对应的运费（人民币）
    """

    weight_g: float
    shipping_cost_rmb: float


class ShippingTable:
    """重量→运费查找表。

    用法:
        table = ShippingTable([
            WeightShippingRow(80, 26.0),
            WeightShippingRow(100, 26.5),
            ...
        ])
        cost = table.lookup(150)   # 线性插值
        cost = table.lookup(150, method="ceil")  # 只向上取整
    """

    def __init__(self, rows: Sequence[WeightShippingRow]):
        if not rows:
            raise ValueError("运费表不能为空")
        sorted_rows = sorted(rows, key=lambda r: r.weight_g)
        self._weights = [r.weight_g for r in sorted_rows]
        self._costs = [r.shipping_cost_rmb for r in sorted_rows]

        # 检查单调性（仅 debug 级别 — 源数据可能有非单调修正）
        for i in range(1, len(self._costs)):
            if self._costs[i] < self._costs[i - 1]:
                logger.debug(
                    "运费表非单调: weight=%g cost=%g -> weight=%g cost=%g",
                    self._weights[i - 1], self._costs[i - 1],
                    self._weights[i], self._costs[i],
                )

    def lookup(self, weight_g: float, method: str = "linear") -> float:
        """根据重量查找运费。

        Args:
            weight_g: 包裹重量（克）。
            method:
                - "linear": 区间线性插值
                - "ceil":   取 ≥ 输入重量的第一个档位值（向上找档）
                - "floor":  取 ≤ 输入重量的第一个档位值（向下找档）

        Returns:
            运费（人民币）。
        """
        if method == "ceil":
            for w, c in zip(self._weights, self._costs):
                if w >= weight_g:
                    return c
            return self._costs[-1]

        if method == "floor":
            result = self._costs[0]
            for w, c in zip(self._weights, self._costs):
                if w <= weight_g:
                    result = c
                else:
                    break
            return result

        # linear interpolation
        if weight_g <= self._weights[0]:
            return self._costs[0]
        if weight_g >= self._weights[-1]:
            return self._costs[-1]

        for i in range(len(self._weights) - 1):
            if self._weights[i] <= weight_g <= self._weights[i + 1]:
                if self._weights[i + 1] == self._weights[i]:
                    return self._costs[i]
                ratio = (weight_g - self._weights[i]) / (
                    self._weights[i + 1] - self._weights[i]
                )
                return self._costs[i] + ratio * (
                    self._costs[i + 1] - self._costs[i]
                )

        # fallback（不应到达）
        return self._costs[-1]

    @classmethod
    def from_excel(
        cls,
        filepath: str,
        station: Station,
        weight_row: int = 9,
        cost_row: int = 10,
    ) -> "ShippingTable":
        """从利润计算 Excel 文件中读取指定站点的运费表。

        Args:
            filepath: Excel 文件路径（如 data/lr.xlsx）。
            station: 站点。
            weight_row: 重量（g）所在行号（1-based，相对于 sheet）。
            cost_row: 运费（元）所在行号。

        Returns:
            ShippingTable 实例。
        """
        try:
            df = pd.read_excel(
                filepath,
                sheet_name=station.value,
                header=None,
            )
        except Exception as e:
            raise ValueError(f"读取 Excel 运费表失败 [{filepath}]: {e}") from e

        # 使用 0-based 索引
        weights = df.iloc[weight_row - 1, 1:].dropna().tolist()
        costs = df.iloc[cost_row - 1, 1:].dropna().tolist()

        # 对齐长度
        min_len = min(len(weights), len(costs))
        rows = [
            WeightShippingRow(float(weights[i]), float(costs[i]))
            for i in range(min_len)
        ]
        return cls(rows)


# ─────────────── 内置运费表（来自 data/lr.xlsx 实测数据）───────────────

# DE 运费表
DE_SHIPPING_TABLE = ShippingTable([
    WeightShippingRow(80, 26.0),
    WeightShippingRow(100, 26.5),
    WeightShippingRow(120, 26.5),
    WeightShippingRow(140, 28.0),
    WeightShippingRow(160, 29.0),
    WeightShippingRow(180, 30.0),
    WeightShippingRow(200, 30.5),
    WeightShippingRow(220, 31.5),
    WeightShippingRow(240, 32.5),
    WeightShippingRow(260, 33.5),
    WeightShippingRow(280, 34.5),
    WeightShippingRow(300, 35.5),
    WeightShippingRow(320, 37.5),
    WeightShippingRow(340, 38.5),
    WeightShippingRow(360, 39.5),
    WeightShippingRow(380, 40.5),
    WeightShippingRow(400, 41.5),
])

# US 运费表
US_SHIPPING_TABLE = ShippingTable([
    WeightShippingRow(80, 31.0),
    WeightShippingRow(100, 31.0),
    WeightShippingRow(120, 32.0),
    WeightShippingRow(140, 34.0),
    WeightShippingRow(160, 35.5),
    WeightShippingRow(180, 37.0),
    WeightShippingRow(200, 36.5),
    WeightShippingRow(220, 38.0),
    WeightShippingRow(240, 39.5),
    WeightShippingRow(260, 41.0),
    WeightShippingRow(280, 42.5),
    WeightShippingRow(300, 44.5),
    WeightShippingRow(320, 46.0),
    WeightShippingRow(340, 47.5),
    WeightShippingRow(360, 49.0),
    WeightShippingRow(380, 51.0),
    WeightShippingRow(400, 52.5),
])

# CA 运费表
CA_SHIPPING_TABLE = ShippingTable([
    WeightShippingRow(80, 21.0),
    WeightShippingRow(100, 22.0),
    WeightShippingRow(120, 23.5),
    WeightShippingRow(140, 24.5),
    WeightShippingRow(160, 26.5),
    WeightShippingRow(180, 27.5),
    WeightShippingRow(200, 28.5),
    WeightShippingRow(220, 30.0),
    WeightShippingRow(240, 30.0),
    WeightShippingRow(260, 31.0),
    WeightShippingRow(280, 32.0),
    WeightShippingRow(300, 33.0),
    WeightShippingRow(320, 35.0),
    WeightShippingRow(340, 36.0),
    WeightShippingRow(360, 37.0),
    WeightShippingRow(380, 38.5),
    WeightShippingRow(400, 39.5),
])

# 内置运费表映射
_BUILTIN_SHIPPING: Dict[Station, ShippingTable] = {
    Station.US: US_SHIPPING_TABLE,
    Station.CA: CA_SHIPPING_TABLE,
    Station.DE: DE_SHIPPING_TABLE,
}


# ==================== 核心计算器 ====================


class ProfitCalculator:
    """多站点利润计算器。

    核心公式:
        毛利率 = (总售价外币 × 汇率 - 成本 - 运费 - 总售价外币 × 汇率 × 扣点率)
                 / (总售价外币 × 汇率)

    即:
        margin = (total_foreign - cost/exchange - ship/exchange
                  - total_foreign * deduction) / total_foreign

    特性:
        - 按售价区间自动匹配扣点率
        - 支持自定义扣点率（覆盖内置配置）
        - 支持重量→运费自动查找
        - 支持 DataFrame 批量计算
        - 支持扣点明细拆解

    用法:
        calc = ProfitCalculator(station=Station.DE)

        # 基础用法
        r = calc.calculate(49.99, 6.99, 7.8, 130.0, 43.5)
        print(f"{r.margin_rate:.2%}")  # → 23.46%

        # 从重量自动查运费
        r = calc.calculate(49.99, 6.99, 7.8, 130.0, weight_g=160)

        # 自定义扣点率
        calc.set_custom_tiers([DeductionTier(None, 0.40, "自定义40%")])

        # 批量计算
        df_results = calc.calculate_batch(dataframe)
    """

    def __init__(
        self,
        station: Station = Station.DE,
        shipping_table: Optional[ShippingTable] = None,
        use_detail: bool = True,
    ):
        """初始化计算器。

        Args:
            station: 目标站点。
            shipping_table: 运费查找表，若为 None 则使用内置表。
            use_detail: 是否在结果中包含扣点明细拆解。
        """
        self.station = station
        self._tiers = StationConfig.get_tiers(station)
        self._shipping_table = shipping_table or _BUILTIN_SHIPPING.get(station)
        self._use_detail = use_detail

    # ── 属性 ──

    @property
    def tiers(self) -> List[DeductionTier]:
        """当前生效的扣点率档位。"""
        return list(self._tiers)

    @property
    def shipping_table(self) -> Optional[ShippingTable]:
        """当前使用的运费查找表。"""
        return self._shipping_table

    # ── 配置方法 ──

    def set_custom_tiers(self, tiers: List[DeductionTier]) -> None:
        """覆盖内置扣点率档位（用于自定义场景或 A/B 测试）。

        Args:
            tiers: 自定义档位列表，按 max_price 升序排列。
        """
        if not tiers:
            raise ValueError("扣点率档位不能为空")
        sorted_tiers = sorted(
            tiers,
            key=lambda t: t.max_price if t.max_price is not None else float("inf"),
        )
        self._tiers = sorted_tiers
        logger.info(
            "自定义扣点率已生效: %s",
            [(t.label, t.rate) for t in sorted_tiers],
        )

    def set_shipping_table(self, table: ShippingTable) -> None:
        """设置运费查找表。"""
        self._shipping_table = table

    # ── 扣点率匹配 ──

    def _find_deduction_rate(self, item_price_foreign: float) -> DeductionTier:
        """根据商品售价匹配综合扣点率档位。

        注意: 匹配的是 item_price（不含运费的售价），而非 total_price。
        这与 Excel 公式一致——Excel 中用 B6（售价）而非 B2（总售价 = B6+B7）
        来判断扣点率档位。但实际上 B2 和 B6 在 15€ 两端的判断几乎相同。

        Args:
            item_price_foreign: 商品售价（外币，不含运费）。

        Returns:
            匹配的 DeductionTier。
        """
        for tier in self._tiers:
            if tier.max_price is None or item_price_foreign < tier.max_price:
                logger.debug(
                    "售价 %.2f 命中档位 %s (扣点率 %.1f%%)",
                    item_price_foreign,
                    tier.label,
                    tier.rate * 100,
                )
                return tier

        # fallback: 使用最后一档
        last = self._tiers[-1]
        logger.debug("售价 %.2f 命中末档 %s", item_price_foreign, last.label)
        return last

    # ── 核心计算 ──

    def _resolve_shipping_cost(
        self,
        weight_g: Optional[float],
        shipping_method: str,
        shipping_cost_rmb: float,
    ) -> float:
        """根据重量查运费表，未提供重量时返回显式传入的运费成本。

        将 weight_g → shipping_cost_rmb 的查表逻辑集中在此处，
        calculate() 和 breakeven_price() 共用，避免重复和分支差异。
        """
        if weight_g is not None:
            if self._shipping_table is None:
                raise ValueError("未设置运费查找表，无法根据重量查运费")
            return self._shipping_table.lookup(weight_g, method=shipping_method)
        return shipping_cost_rmb

    def calculate(
        self,
        item_price: float,
        shipping_price: float = 0.0,
        exchange_rate: float = 7.8,
        sku_cost_rmb: float = 0.0,
        shipping_cost_rmb: float = 0.0,
        *,
        weight_g: Optional[float] = None,
        shipping_method: str = "ceil",
    ) -> ProfitResult:
        """计算单条商品利润。

        Args:
            item_price: 商品售价（外币，如 €49.99）。
            shipping_price: 向买家收取的运费（外币，如 €6.99）。
            exchange_rate: 汇率（本币/外币，如 7.8 表示 1€ = 7.8¥）。
            sku_cost_rmb: SKU 成本（人民币）。
            shipping_cost_rmb: 运费成本（人民币）。若提供 weight_g 则忽略此值。
            weight_g: 包裹重量（克），若提供则从运费表自动查找 shipping_cost_rmb。
            shipping_method: 查表方法 ("linear" | "ceil" | "floor")，默认 ceil。

        Returns:
            ProfitResult 包含毛利、净利、扣点明细等。
        """
        # ── 输入校验 ──
        if item_price <= 0:
            raise ValueError(f"item_price 必须 > 0，当前值: {item_price}")
        if exchange_rate <= 0:
            raise ValueError(f"exchange_rate 必须 > 0，当前值: {exchange_rate}")

        # ── 如果给定了重量，从查表获取运费 ──
        shipping_cost_rmb = self._resolve_shipping_cost(
            weight_g, shipping_method, shipping_cost_rmb
        )

        # ── 总售价 ──
        total_foreign = item_price + shipping_price
        total_rmb = total_foreign * exchange_rate

        # ── 匹配扣点率 ──
        tier = self._find_deduction_rate(item_price)

        # ── 综合扣除金额 ──
        deduction_amount_rmb = total_rmb * tier.rate

        # ── 总成本 ──
        cost_total_rmb = sku_cost_rmb + shipping_cost_rmb

        # ── 净利润 ──
        profit_rmb = total_rmb - cost_total_rmb - deduction_amount_rmb

        # ── 毛利率 ──
        margin_rate = profit_rmb / total_rmb if total_rmb != 0 else 0.0

        # ── 可选的扣点明细 ──
        detail = None
        if self._use_detail:
            raw_detail = StationConfig.get_detail(self.station, tier.label)
            if raw_detail:
                detail = {
                    k: total_rmb * v for k, v in raw_detail.items()
                }

        return ProfitResult(
            station=self.station,
            item_price=item_price,
            shipping_price=shipping_price,
            exchange_rate=exchange_rate,
            sku_cost_rmb=sku_cost_rmb,
            shipping_cost_rmb=shipping_cost_rmb,
            total_price_foreign=total_foreign,
            total_price_rmb=total_rmb,
            deduction_rate=tier.rate,
            deduction_tier_label=tier.label,
            deduction_amount_rmb=deduction_amount_rmb,
            cost_rmb_total=cost_total_rmb,
            profit_rmb=profit_rmb,
            margin_rate=margin_rate,
            detail=detail,
        )

    # ── 批量计算 ──

    def calculate_batch(
        self,
        df: pd.DataFrame,
        *,
        item_price_col: str = "item_price",
        shipping_price_col: str = "shipping_price",
        exchange_rate_col: str = "exchange_rate",
        sku_cost_col: str = "sku_cost_rmb",
        shipping_cost_col: str = "shipping_cost_rmb",
        weight_col: Optional[str] = "weight_g",
        shipping_method: str = "ceil",
    ) -> pd.DataFrame:
        """批量计算利润，返回带结果列的 DataFrame。

        输入的 DataFrame 需包含以下列（可通过参数自定义列名）：
            - item_price
            - shipping_price
            - exchange_rate
            - sku_cost_rmb
            - shipping_cost_rmb
            - weight_g（可选）

        返回的 DataFrame 在原基础上追加以下列：
            - total_price_foreign
            - total_price_rmb
            - deduction_rate
            - deduction_tier
            - profit_rmb
            - margin_rate

        Args:
            df: 输入 DataFrame。
            item_price_col: 商品售价列名。
            shipping_price_col: 运费收入列名。
            exchange_rate_col: 汇率列名。
            sku_cost_col: SKU 成本列名。
            shipping_cost_col: 运费成本列名。
            weight_col: 重量列名（可选）。
            shipping_method: 查表方法。

        Returns:
            追加了结果列的 DataFrame（拷贝，不修改原 df）。
        """
        result_df = df.copy()

        for idx, row in result_df.iterrows():
            # 构建 kwargs
            kwargs: Dict[str, Any] = {
                "item_price": float(row.get(item_price_col, 0) or 0),
                "shipping_price": float(row.get(shipping_price_col, 0) or 0),
                "exchange_rate": float(row.get(exchange_rate_col, 7.8) or 7.8),
                "sku_cost_rmb": float(row.get(sku_cost_col, 0) or 0),
                "shipping_method": shipping_method,
            }

            # 优先用重量查运费，否则用显式运费成本
            if weight_col and weight_col in row and pd.notna(row[weight_col]):
                kwargs["weight_g"] = float(row[weight_col])
            else:
                kwargs["shipping_cost_rmb"] = float(
                    row.get(shipping_cost_col, 0) or 0
                )

            try:
                r = self.calculate(**kwargs)
            except ValueError as e:
                logger.warning("第 %d 行计算失败: %s", idx, e)
                continue

            result_df.loc[idx, "total_price_foreign"] = r.total_price_foreign
            result_df.loc[idx, "total_price_rmb"] = r.total_price_rmb
            result_df.loc[idx, "deduction_rate"] = r.deduction_rate
            result_df.loc[idx, "deduction_tier"] = r.deduction_tier_label
            result_df.loc[idx, "profit_rmb"] = round(r.profit_rmb, 2)
            result_df.loc[idx, "margin_rate"] = round(r.margin_rate, 4)
            result_df.loc[idx, "cost_rmb_total"] = round(r.cost_rmb_total, 2)
            result_df.loc[idx, "deduction_amount_rmb"] = round(
                r.deduction_amount_rmb, 2
            )

        return result_df

    # ── 便捷方法 ──

    def calculate_from_weight(
        self,
        item_price: float,
        weight_g: float,
        shipping_price: float = 0.0,
        exchange_rate: float = 7.8,
        sku_cost_rmb: float = 0.0,
        shipping_method: str = "ceil",
    ) -> ProfitResult:
        """根据重量自动查运费并计算利润（便捷方法）。"""
        return self.calculate(
            item_price=item_price,
            shipping_price=shipping_price,
            exchange_rate=exchange_rate,
            sku_cost_rmb=sku_cost_rmb,
            weight_g=weight_g,
            shipping_method=shipping_method,
        )

    def roi(self, **kwargs: Any) -> float:
        """计算投入产出比 ROI = 净利润 / 总成本。

        参数同 calculate() 方法。

        Returns:
            ROI 值（0.5 表示 50% 回报率）。
        """
        result = self.calculate(**kwargs)
        if result.cost_rmb_total == 0:
            return float("inf")
        return result.profit_rmb / result.cost_rmb_total

    def breakeven_price(
        self,
        exchange_rate: float = 7.8,
        sku_cost_rmb: float = 0.0,
        shipping_cost_rmb: float = 0.0,
        shipping_price: float = 0.0,
        target_margin: float = 0.15,
        *,
        weight_g: Optional[float] = None,
        shipping_method: str = "ceil",
    ) -> float:
        """反算：给定目标毛利率，求最低售价。

        由公式: margin = (P·R - C - S - P·R·k) / (P·R)
        求解 P（外币售价）:
            P = (C + S) / (R · (1 - k - margin))

        其中:
            P = item_price（商品售价，外币）
            R = exchange_rate（汇率）
            C = sku_cost_rmb
            S = shipping_cost_rmb
            k = deduction_rate（综合扣点率）

        Args:
            exchange_rate: 汇率。
            sku_cost_rmb: SKU 成本（人民币）。
            shipping_cost_rmb: 运费成本（人民币）。
            shipping_price: 向买家收取的运费（外币）。
            target_margin: 目标毛利率（0.15 = 15%）。
            weight_g: 若提供则从运费表查运费。
            shipping_method: 查表方法 ("linear" | "ceil" | "floor")，默认 ceil。

        Returns:
            最低售价（外币，不含运费）。
        """
        if weight_g is not None and self._shipping_table is not None:
            shipping_cost_rmb = self._resolve_shipping_cost(
                weight_g, shipping_method, shipping_cost_rmb
            )

        # 需要知道扣点率，但扣点率又取决于售价——这里用迭代法
        # 先用中间档扣点率估算
        guess_rate = self._tiers[len(self._tiers) // 2].rate

        for _ in range(20):  # 最多迭代 20 次
            denominator = exchange_rate * (1 - guess_rate - target_margin)
            if denominator <= 0:
                raise ValueError(
                    f"目标毛利率 {target_margin:.1%} 加扣点率 {guess_rate:.1%} "
                    f"超过 100%，无法正算价格"
                )
            p = (sku_cost_rmb + shipping_cost_rmb) / denominator

            # 重新匹配扣点率
            new_tier = self._find_deduction_rate(p)
            if abs(new_tier.rate - guess_rate) < 1e-9:
                break
            guess_rate = new_tier.rate

        # 减去运费收入得到纯售价
        return p - shipping_price

    def summary(self, result: ProfitResult) -> str:
        """生成单次计算结果的人类可读摘要。"""
        Y = "RMB"  # 避免 Windows GBK 终端编码问题
        lines = [
            f"{'='*40}",
            f"  {self.station.value} Profit Analysis",
            f"{'='*40}",
            f"Station:        {self.station.value}",
            f"Item Price:     {result.item_price:.2f}",
            f"Shipping Income:{result.shipping_price:.2f}",
            f"Total (Foreign):{result.total_price_foreign:.2f}",
            f"FX Rate:        {result.exchange_rate:.2f}",
            f"Total ({Y}):    {result.total_price_rmb:.2f}",
            f"SKU Cost:       {result.sku_cost_rmb:.2f}",
            f"Ship Cost:      {result.shipping_cost_rmb:.2f}",
            f"Total Cost:     {result.cost_rmb_total:.2f}",
            f"Deduction Tier: {result.deduction_tier_label} "
            f"({result.deduction_rate:.1%})",
            f"Deduction Amt:  {result.deduction_amount_rmb:.2f}",
            f"{'-'*40}",
            f"Net Profit:     {result.profit_rmb:.2f} {Y}",
            f"Margin Rate:    {result.margin_rate:.2%}",
        ]
        if result.detail:
            lines.append("")
            lines.append("Deduction Detail:")
            for name, amount in result.detail.items():
                lines.append(f"  {name:<10s} {amount:.2f} {Y}")
        return "\n".join(lines)


# ==================== 便捷函数 ====================

_DEFAULT_RATES: Dict[Station, float] = {
    Station.DE: 7.8,
    Station.US: 7.0,
}


def quick_profit(
    station: Station,
    item_price: float,
    sku_cost_rmb: float,
    weight_g: float,
    shipping_price: float = 0.0,
    exchange_rate: Optional[float] = None,
) -> ProfitResult:
    """快速利润计算（从重量自动查运费）。

    Args:
        station: 目标站点。
        item_price: 商品售价（外币）。
        sku_cost_rmb: SKU 成本（人民币）。
        weight_g: 包裹重量（克）。
        shipping_price: 向买家收取的运费（外币）。
        exchange_rate: 汇率（默认：DE=7.8, US=7.0）。

    Returns:
        ProfitResult。
    """
    if exchange_rate is None:
        exchange_rate = _DEFAULT_RATES.get(station, 7.8)
    calc = ProfitCalculator(station)
    return calc.calculate_from_weight(
        item_price=item_price,
        weight_g=weight_g,
        shipping_price=shipping_price,
        exchange_rate=exchange_rate,
        sku_cost_rmb=sku_cost_rmb,
    )


# 保留旧函数名（向后兼容别名）
def quick_de(
    item_price: float,
    sku_cost_rmb: float,
    weight_g: float,
    shipping_price: float = 0.0,
    exchange_rate: float = 7.8,
) -> ProfitResult:
    """德国站 FBM 快速计算（向后兼容，推荐使用 quick_profit）。"""
    return quick_profit(
        Station.DE, item_price, sku_cost_rmb, weight_g,
        shipping_price=shipping_price, exchange_rate=exchange_rate,
    )


def quick_us(
    item_price: float,
    sku_cost_rmb: float,
    weight_g: float,
    shipping_price: float = 0.0,
    exchange_rate: float = 7.0,
) -> ProfitResult:
    """美国站快速计算（向后兼容，推荐使用 quick_profit）。"""
    return quick_profit(
        Station.US, item_price, sku_cost_rmb, weight_g,
        shipping_price=shipping_price, exchange_rate=exchange_rate,
    )


# ==================== 自测入口 ====================

if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    print("\n" + "=" * 50)
    print("  DE FBM Profit Calculation Verification")
    print("=" * 50)

    calc = ProfitCalculator(Station.DE)

    # 验证: 售价 49.99 > 15 EUR → 命中 44.5% 档位
    # Excel Row 2 (37.5%): 0.234625
    # Excel Row 3 (44.5%): 0.164625  ← 代码应匹配此项
    result_high = calc.calculate(
        item_price=49.99,
        shipping_price=6.99,
        exchange_rate=7.8,
        sku_cost_rmb=130.0,
        shipping_cost_rmb=43.5,
    )

    print(calc.summary(result_high))

    expected_high = 0.164625  # 15 EUR+ formula
    print(f"\nExcel Row3 (44.5%): {expected_high:.4%}")
    print(f"Code calculated:     {result_high.margin_rate:.4%}")
    delta = abs(result_high.margin_rate - expected_high)
    print(f"Delta:               {delta:.6f}")
    assert delta < 1e-5, f"High-tier margin mismatch: {delta}"

    # 验证: 售价 8.99 < 15 EUR → 命中 37.5% 档位
    # 直接用 Excel 中 Row 2 的公式预期: 0.316279
    result_low = calc.calculate(
        item_price=8.99,
        shipping_price=5.96,
        exchange_rate=7.8,
        sku_cost_rmb=10.0,
        shipping_cost_rmb=26.0,
    )
    expected_low = 0.316279  # 0-15 EUR formula
    delta_low = abs(result_low.margin_rate - expected_low)
    assert delta_low < 1e-5, f"Low-tier margin mismatch: {delta_low}"
    print(f"\n0-15 EUR verify: expected={expected_low:.4%}, "
          f"got={result_low.margin_rate:.4%}, delta={delta_low:.6f} [PASS]")

    # 重量查运费验证
    print("\n--- Weight-based Shipping Lookup ---")
    result_w = calc.calculate_from_weight(
        item_price=49.99,
        weight_g=160,
        exchange_rate=7.8,
        sku_cost_rmb=130.0,
    )
    print(f"weight=160g -> shipping=RMB {result_w.shipping_cost_rmb:.2f}, "
          f"margin={result_w.margin_rate:.2%}")

    # 反算最低售价
    print("\n--- Breakeven Price (target margin 20%) ---")
    min_price = calc.breakeven_price(
        exchange_rate=7.8,
        sku_cost_rmb=130.0,
        shipping_cost_rmb=43.5,
        shipping_price=6.99,
        target_margin=0.20,
    )
    print(f"Minimum item price: EUR {min_price:.2f}")

    print("\nAll tests passed.")
