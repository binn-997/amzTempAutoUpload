import pandas as pd

from amazon_listing_toolkit.color_translator import ColorTranslator


def test_translates_and_labels_duplicate_variants() -> None:
    frame = pd.DataFrame([
        ["", "child-a", "parent", "", "", "", "Blue", "M"],
        ["", "child-red", "parent", "", "", "", "Red", "L"],
        ["", "child-b", "parent", "", "", "", "Blue", "M"],
    ])

    result = ColorTranslator().process(frame, color_col=6, size_col=7, parent_sku_col=2)

    assert result.index_to_new_color == {0: "BlauV1", 1: "Rot", 2: "BlauV2"}
