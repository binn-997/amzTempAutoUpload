from pathlib import Path

import pytest

from amazon_listing_toolkit.config_loader import (
    ConfigError,
    load_config,
    resolve_category_config,
    validate_config,
)


def test_default_config_resolves_category() -> None:
    config_path = Path(__file__).parents[1] / "src" / "amazon_listing_toolkit" / "config" / "categories.yaml"
    config = load_config(str(config_path))

    category = resolve_category_config(config, "hd")

    assert category["size_class_col"] == "BA"
    assert category["size_target_cols"] == ["BB", "CH", "FC"]


def test_invalid_config_is_rejected() -> None:
    with pytest.raises(ConfigError, match="必填项"):
        validate_config({"categories": {}})
