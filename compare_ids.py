import pandas as pd
from typing import Any, Dict, List, Optional, Tuple

def compare_series(
    a: pd.Series,
    b: pd.Series,
    name_a: str = "A",
    name_b: str = "B",
    dropna_for_sets: bool = True,
) -> Dict[str, Any]:
    """
    Compare two pandas Series and return a structured summary:
      - intersection (overlap)
      - only_in_a, only_in_b
      - union, symmetric_difference
      - counts summary (including NaNs)
      - value_counts for each series
      - details DataFrame with per-value counts and membership flags

    Parameters:
      a, b: pandas Series to compare
      name_a, name_b: labels used in outputs
      dropna_for_sets: if True, NaNs are excluded from set-based comparisons
                       (but NaN counts are still reported separately)

    Returns:
      dict with keys:
        - intersection: list
        - only_in_a: list
        - only_in_b: list
        - union: list
        - symmetric_difference: list
        - counts: dict with element and NaN counts
        - value_counts_a: pd.Series
        - value_counts_b: pd.Series
        - details: pd.DataFrame with columns:
            [count_in_A, count_in_B, in_A, in_B, in_both, only_in_A, only_in_B]
          indexed by distinct values observed across both series, ordered by first appearance
    """
    if not isinstance(a, pd.Series) or not isinstance(b, pd.Series):
        raise TypeError("Both inputs must be pandas Series.")

    # Capture NaN counts (treated separately)
    na_count_a = int(a.isna().sum())
    na_count_b = int(b.isna().sum())

    # Series with or without NaNs for set-like operations
    as_set = a.dropna() if dropna_for_sets else a
    bs_set = b.dropna() if dropna_for_sets else b

    # Build order of first appearance across both series to make result readable
    def first_appearance_order(s: pd.Series) -> Dict[Any, int]:
        # Use dict to keep first index per value
        first_idx = {}
        for i, v in enumerate(s.tolist()):
            if v not in first_idx:
                first_idx[v] = i
        return first_idx

    order_a = first_appearance_order(as_set)
    order_b = first_appearance_order(bs_set)
    # Combine orders: position in concatenated appearance
    combined_order = {}
    next_pos = 0
    for source in (order_a, order_b):
        for k in source:
            if k not in combined_order:
                combined_order[k] = next_pos
                next_pos += 1

    # Compute sets
    set_a = set(as_set.tolist())
    set_b = set(bs_set.tolist())
    inter = set_a & set_b
    only_a = set_a - set_b
    only_b = set_b - set_a
    uni = set_a | set_b
    symdiff = set_a ^ set_b

    # Value counts (include NaNs here for completeness)
    vc_a = a.value_counts(dropna=False)
    vc_b = b.value_counts(dropna=False)

    # Details DataFrame indexed by union values, ordered by first appearance
    union_vals_sorted = sorted(list(uni), key=lambda x: combined_order.get(x, float("inf")))
    details = pd.DataFrame(
        {
            f"count_in_{name_a}": [int(vc_a.get(v, 0)) for v in union_vals_sorted],
            f"count_in_{name_b}": [int(vc_b.get(v, 0)) for v in union_vals_sorted],
        },
        index=union_vals_sorted,
    )
    details[f"in_{name_a}"] = details[f"count_in_{name_a}"] > 0
    details[f"in_{name_b}"] = details[f"count_in_{name_b}"] > 0
    details["in_both"] = details[f"in_{name_a}"] & details[f"in_{name_b}"]
    details[f"only_in_{name_a}"] = details[f"in_{name_a}"] & (~details[f"in_{name_b}"])
    details[f"only_in_{name_b}"] = details[f"in_{name_b}"] & (~details[f"in_{name_a}"])

    # Assemble counts summary
    counts_summary = {
        f"n_{name_a}": int(len(a)),
        f"n_{name_b}": int(len(b)),
        f"n_unique_{name_a}": int(a.dropna().nunique()),
        f"n_unique_{name_b}": int(b.dropna().nunique()),
        f"n_overlap_values": int(len(inter)),
        f"n_only_in_{name_a}": int(len(only_a)),
        f"n_only_in_{name_b}": int(len(only_b)),
        "n_union_values": int(len(uni)),
        "n_symmetric_difference": int(len(symdiff)),
        f"na_count_{name_a}": na_count_a,
        f"na_count_{name_b}": na_count_b,
    }

    result = {
        "intersection": _ordered_list(inter, combined_order),
        "only_in_a": _ordered_list(only_a, combined_order),
        "only_in_b": _ordered_list(only_b, combined_order),
        "union": _ordered_list(uni, combined_order),
        "symmetric_difference": _ordered_list(symdiff, combined_order),
        "counts": counts_summary,
        "value_counts_a": vc_a,
        "value_counts_b": vc_b,
        "details": details,
    }
    return result


def _ordered_list(values: set, order_map: Dict[Any, int]) -> List[Any]:
    """Helper to return a set as a list ordered by first appearance."""
    return sorted(list(values), key=lambda x: order_map.get(x, float("inf")))

if __name__ == '__main__':
    # Example usage:
    s1 = pd.Series(["x", "y", "y", None, "z"])
    s2 = pd.Series(["y", "z", "z", "w"])
    res = compare_series(s1, s2, name_a="A", name_b="B")
    print(res["counts"])
    print(res["intersection"])
    print(res["only_in_a"])
    print(res["details"])