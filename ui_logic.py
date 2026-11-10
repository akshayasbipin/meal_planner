"""Pure helpers for app.py (no Streamlit import, so they can be tested on their own)."""
from __future__ import annotations

from collections import defaultdict

import pandas as pd

import meal_planner as mp

# ---------------------------------------------------------------- pantry
def pantry_df(wb) -> pd.DataFrame:
    ws = mp.ensure_pantry_sheet(wb)
    headers = [c.value for c in ws[1]]
    ii, qi, ui = (mp.find_col(headers, ["Item"]), mp.find_col(headers, ["Available quantity", "Quantity", "Qty"]),
                  mp.find_col(headers, ["Unit"]))
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if ii is not None and r[ii]:
            q = r[qi] if qi is not None else None
            try:
                q = float(q) if q not in (None, "") else None
            except (TypeError, ValueError):
                q = None
            rows.append({"Item": str(r[ii]).strip(), "Available quantity": q,
                         "Unit": str(r[ui] or "") if ui is not None else ""})
    return pd.DataFrame(rows, columns=["Item", "Available quantity", "Unit"])


def df_to_pantry_rows(df: pd.DataFrame) -> list[dict]:
    out = []
    for _, r in df.iterrows():
        item = str(r.get("Item") or "").strip()
        if not item or item.lower() == "nan":
            continue
        q = r.get("Available quantity")
        q = None if pd.isna(q) else (int(q) if float(q).is_integer() else float(q))
        unit = r.get("Unit")
        out.append({"item": item, "qty": q, "unit": "" if pd.isna(unit) else str(unit)})
    return out


def merge_bought(pantry: list[dict], bought: list[dict], aliases: dict):
    """Add newly bought quantities to the pantry. Returns (new_rows, messages)."""
    rows = [dict(r) for r in pantry]
    msgs = []
    for b in bought:
        key = mp.item_key(b["item"], aliases)
        hit = next((r for r in rows if mp.item_key(r["item"], aliases) == key), None)
        qty = b.get("qty")
        if hit is None:
            rows.append({"item": b["item"], "qty": qty, "unit": b.get("unit", "")})
        elif qty in (None, ""):
            msgs.append(f"{b['item']}: no quantity entered, pantry unchanged.")
        elif hit["qty"] in (None, "") or (hit["unit"] or "").lower() != (b.get("unit") or "").lower():
            if hit["qty"] not in (None, ""):
                msgs.append(f"{b['item']}: units differ ({hit['unit']} vs {b.get('unit')}), replaced the old amount.")
            hit["qty"], hit["unit"] = qty, b.get("unit", "")
        else:
            try:
                total = float(hit["qty"]) + float(qty)
            except (TypeError, ValueError):
                hit["qty"], hit["unit"] = qty, b.get("unit", "")
            else:
                hit["qty"] = int(total) if total.is_integer() else total
    return rows, msgs


def meal_options(names: list[str], current: str, other_meal: str, adjacent: set[str]) -> list[str]:
    eligible = [name for name in names if name != other_meal and name not in adjacent]
    return sorted(set(eligible + [current]), key=str.lower)


# ---------------------------------------------------------------- shopping
SHOP_COLS = ["Item", "Category", "Status", "Needed for", "Meals", "In pantry", "Notes"]


def shopping_df(rows: list[list]) -> pd.DataFrame:
    return pd.DataFrame([r[:7] for r in rows], columns=SHOP_COLS)


def shopping_text(rows: list[list], title: str) -> str:
    groups = defaultdict(list)
    for item, cat, status, uses, *_ in rows:
        if status != "In stock":
            groups[cat].append(f"- {item}" + (f" ({uses})" if cat != "Not in Grocery sheet" and len(uses) < 40 else ""))
    lines = [title, ""]
    for cat in sorted(groups, key=lambda c: (c == "Not in Grocery sheet", c)):
        label = "Check at home / optional" if cat == "Not in Grocery sheet" else cat
        lines += [f"{label}:"] + groups[cat] + [""]
    return "\n".join(lines).strip()