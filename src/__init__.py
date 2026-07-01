"""src 包初始化。"""

from src.processors.base_processor import BaseProcessor, ProcessingError, IOFailure
from src.profit_calculator import (
    ProfitCalculator,
    ProfitResult,
    DeductionTier,
    Station,
    StationConfig,
    ShippingTable,
    WeightShippingRow,
    quick_de,
    quick_us,
)
from src.color_translator import (
    ColorTranslator,
    ColorGroup,
    TranslateResult,
    COLOR_TRANSLATIONS,
    quick_translate,
)

__all__ = [
    "BaseProcessor", "ProcessingError", "IOFailure",
    "ProfitCalculator", "ProfitResult", "DeductionTier",
    "Station", "StationConfig",
    "ShippingTable", "WeightShippingRow",
    "quick_de", "quick_us",
    "ColorTranslator", "ColorGroup", "TranslateResult",
    "COLOR_TRANSLATIONS", "quick_translate",
]
