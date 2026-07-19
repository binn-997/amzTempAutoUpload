import pandas as pd

from amazon_listing_toolkit.processors.base_processor import BaseProcessor, sanitize_cell_value


def test_preprocess_inherits_parent_content_and_normalizes_size() -> None:
    processor = BaseProcessor({"size_replacements": {r"\bXXXL\b": "3XL"}})
    parent = [""] * 60
    parent[1], parent[2], parent[3] = "PARENT", "PARENT", "Title"
    parent[40:45] = ["Bullet 1", "Bullet 2", "Bullet 3", "Bullet 4", "Bullet 5"]
    parent[45] = "Keyword"
    child = [""] * 60
    child[1], child[2], child[3], child[6], child[7] = "CHILD", "PARENT", "Title", "Blue", "XXXL"
    frame = pd.DataFrame([parent, child])

    result = processor.preprocess_source_data(frame)

    assert result.loc[1, "size"] == "3XL"
    assert result.loc[1, "final_keywords"] == "Keyword"
    assert result.loc[1, "final_bullet1"] == "Bullet 1"


def test_formula_values_are_escaped() -> None:
    assert sanitize_cell_value("=SUM(A1:A2)") == "'=SUM(A1:A2)"
    assert sanitize_cell_value("ordinary text") == "ordinary text"
