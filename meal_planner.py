"""Excel-backed weekly meal planner with deterministic meal selection."""
from __future__ import annotations

import argparse
import json
import math
import random
import re
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = ROOT / "Grocery.xlsx"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
OUTPUT_SHEETS = ("Meal Plan", "Shopping List", "Meal Prep")

DEFAULT_CONFIG: dict[str, Any] = {
    "cook_for": 1,
    "buy_for": 3,
    "fixed_breakfast": "Chocos with milk + 1 boiled egg",
    "weekly_staples": ["Milk", "Chocos", "Egg"],   # needed for the fixed breakfast
    "can_cook": ["rice"],
    "cannot_cook_yet": ["roti"],
    "serve_rice": True,           # auto-add rice as a side for dishes matching the keywords
    "rice_with": ["dal", "curry", "rajma", "poriyal", "creamy", "soy"],
    "max_repeats_lunch": 2,
    "max_repeats_dinner": 1,
    # words in an ingredient list that you always have / never need to buy
    "ignore_ingredients": ["salt", "oil", "powder", "masala", "jeera", "pepper", "chilli", "bay leaves",
                           "cinamaon", "cinnamon", "nutmeg", "turmeric", "corrainder", "corriander", "coriander",
                           "mirchi", "pasre", "lemon juice", "water", "dal"],
    # spelling/synonym fixes: workbook word -> Grocery sheet item
    "aliases": {"palak": "Spinach", "rajma": "Rajama", "soya chunk": "Soy chunks",
                "soya chunks": "Soy chunks", "soy chunk": "Soy chunks", "almonds": "Almond"},
}


# ----------------------------------------------------------------------------- helpers
def load_config(wb) -> dict:
    cfg = dict(DEFAULT_CONFIG)
    ws = get_sheet(wb, "Planner Settings", required=False)
    if ws is None:
        return cfg
    headers = [cell.value for cell in ws[1]]
    key_i = find_col(headers, ["Setting", "Key"])
    value_i = find_col(headers, ["Value"])
    if key_i is None or value_i is None:
        return cfg
    values = {str(row[key_i]).strip(): row[value_i] for row in ws.iter_rows(min_row=2, values_only=True)
              if key_i < len(row) and value_i < len(row) and row[key_i]}
    for key, default in DEFAULT_CONFIG.items():
        raw = values.get(key)
        if raw is None:
            continue
        if isinstance(default, list):
            cfg[key] = [part.strip() for part in str(raw).split(",") if part.strip()]
        elif isinstance(default, dict):
            cfg[key] = json.loads(str(raw))
        elif isinstance(default, bool):
            cfg[key] = raw if isinstance(raw, bool) else norm(raw) in {"true", "yes", "1"}
        elif isinstance(default, int):
            cfg[key] = int(raw)
        else:
            cfg[key] = str(raw)
    return cfg


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def sheet_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def get_sheet(wb, name: str, required: bool = True):
    """Find a sheet ignoring case, spaces and underscores ('What_can_i_make' == 'What can i make')."""
    for title in wb.sheetnames:
        if sheet_key(title) == sheet_key(name):
            return wb[title]
    if required:
        raise ValueError(f"Missing worksheet '{name}'. Found: {wb.sheetnames}")
    return None


def find_col(headers: list, candidates: list[str]) -> int | None:
    normalized = [norm(h) for h in headers]
    for candidate in candidates:
        if norm(candidate) in normalized:
            return normalized.index(norm(candidate))
    return None


def item_key(name: str, aliases: dict) -> str:
    """Canonical key so 'Eggs' == 'Egg', 'Rajama' == 'rajma', 'Soy chunks' == 'soya chunk'."""
    n = norm(name)
    alias_map = {norm(k): norm(v) for k, v in aliases.items()}
    n = alias_map.get(n, n)
    return n[:-1] if len(n) > 3 and n.endswith("s") else n


# ----------------------------------------------------------------------------- reading the workbook
def read_grocery_candidates(ws) -> list[dict]:
    headers = [c.value for c in ws[1]]
    out = []
    for col, header in enumerate(headers, start=1):
        if not header:
            continue
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row, col).value
            if v is not None and str(v).strip():
                out.append({"item": str(v).strip(), "category": str(header).strip()})
    return out


def read_pantry(ws, aliases: dict) -> dict[str, dict]:
    """Header-aware. Only rows with a non-zero quantity count as 'in stock'."""
    if ws is None:
        return {}
    headers = [c.value for c in ws[1]]
    item_i = find_col(headers, ["Item"])
    qty_i = find_col(headers, ["Available quantity", "Quantity", "Qty"])
    unit_i = find_col(headers, ["Unit"])
    if item_i is None or qty_i is None:
        return {}
    pantry = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        item = row[item_i] if item_i < len(row) else None
        qty = row[qty_i] if qty_i < len(row) else None
        if not item or qty in (None, "", 0, "0"):
            continue
        unit = row[unit_i] if unit_i is not None and unit_i < len(row) else ""
        if isinstance(qty, float) and qty.is_integer():
            qty = int(qty)
        pantry[item_key(item, aliases)] = {"item": str(item).strip(), "qty": f"{qty} {unit or ''}".strip()}
    return pantry


def parse_meals(text: Any) -> set[str]:
    t = norm(text)
    meals = set()
    if re.search(r"\bbreakfast\b|\bb\b", t):
        meals.add("b")
    if re.search(r"\blunch\b|\bl\b", t):
        meals.add("l")
    if re.search(r"\bdinner\b|\bd\b", t):
        meals.add("d")
    return meals  # empty set = not specified (treated as flexible lunch/dinner)


def parse_status(text: Any) -> str:
    t = norm(text)
    if t.startswith("known"):
        return "known"
    if t.startswith("learn"):
        return "learning"
    if t.startswith("not tried"):
        return "not tried"
    return "unknown"


def read_dishes(ws) -> list[dict]:
    headers = [c.value for c in ws[1]]
    idx = {
        "name": find_col(headers, ["Dish", "Recipe", "Name"]),
        "link": find_col(headers, ["Link", "URL"]),
        "ingredients": find_col(headers, ["Ingredients"]),
        "ease": find_col(headers, ["Ease (E/M/D)", "Ease", "Difficulty"]),
        "serving": find_col(headers, ["Serving size (1 person)", "Serving size", "Portion"]),
        "meal": find_col(headers, ["B / L / D", "Meal", "Meal type"]),
        "status": find_col(headers, ["Cooking status", "Familiarity"]),
    }
    if idx["name"] is None:
        raise ValueError("No 'Dish' column found in the dishes sheet.")
    dishes = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        def get(k):
            i = idx[k]
            return row[i] if i is not None and i < len(row) and row[i] is not None else ""
        name = str(get("name")).strip()
        if not name:
            continue
        ease = norm(get("ease"))[:1]  # e / m / d / ""
        dishes.append({
            "name": name, "ingredients": str(get("ingredients")).strip(),
            "ease": {"e": "Easy", "m": "Medium", "d": "Difficult"}.get(ease, "Unknown"),
            "serving": str(get("serving")).strip(),
            "meals": parse_meals(get("meal")), "status": parse_status(get("status")), "row": r,
        })
    return dishes


def read_snacks(ws) -> list[str]:
    return [str(r[0]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0] and str(r[0]).strip()]


def read_nutrition_priorities(ws) -> list[dict]:
    if ws is None:
        return []
    header_row = next(
        ((index, list(row)) for index, row in enumerate(ws.iter_rows(values_only=True), start=1)
         if find_col(list(row), ["Nutrient"]) is not None),
        None,
    )
    if header_row is None:
        return []
    header_index, headers = header_row
    nutrient_i = find_col(headers, ["Nutrient"])
    foods_i = find_col(headers, ["Foods to include", "Foods", "Keywords"])
    target_i = find_col(headers, ["Weekly target servings", "Weekly target", "Target servings"])
    if nutrient_i is None or foods_i is None or target_i is None:
        return []
    priorities = []
    for row in ws.iter_rows(min_row=header_index + 1, values_only=True):
        if max(nutrient_i, foods_i, target_i) >= len(row):
            continue
        nutrient, foods, target = row[nutrient_i], row[foods_i], row[target_i]
        if not nutrient or not foods or target in (None, ""):
            continue
        terms = [norm(term) for term in re.split(r"[,;/]|\s+or\s+", str(foods), flags=re.IGNORECASE) if norm(term)]
        try:
            target_count = int(float(target))
        except (TypeError, ValueError):
            continue
        if terms and target_count > 0:
            priorities.append({"name": str(nutrient).strip(), "foods": terms, "target": target_count})
    return priorities


def ensure_pantry_sheet(wb):
    """Create Pantry only if missing. Never touches or re-seeds an existing one."""
    if get_sheet(wb, "Pantry", required=False) is None:
        ws = wb.create_sheet("Pantry")
        ws.append(["Item", "Available quantity", "Unit"])
    return get_sheet(wb, "Pantry")


# ----------------------------------------------------------------------------- rules in code
def classify_dishes(dishes: list[dict], cfg: dict):
    blocked = [norm(x) for x in cfg["cannot_cook_yet"]]
    sides = {norm(x) for x in cfg["can_cook"]}
    lunch, dinner, excluded = [], [], []
    for d in dishes:
        n = norm(d["name"])
        if any(re.search(rf"\b{re.escape(b)}\b", n) for b in blocked):
            excluded.append((d["name"], "you can't cook it yet")); continue
        if n in sides:
            excluded.append((d["name"], "side dish - added automatically")); continue
        if d["meals"] == {"b"}:
            excluded.append((d["name"], "breakfast only (breakfast is fixed)")); continue
        if (not d["meals"] or "l" in d["meals"]) and d["status"] == "known" and d["ease"] != "Difficult":
            lunch.append(d)
        if not d["meals"] or "d" in d["meals"]:
            dinner.append(d)
        if d not in lunch and d not in dinner:
            why = "Lunch-only but not marked Known" if "l" in d["meals"] else "no usable meal type"
            excluded.append((d["name"], why))
    return lunch, dinner, excluded


def limits(cfg, lunch_pool, dinner_pool) -> dict:
    cap_l = max(cfg["max_repeats_lunch"], math.ceil(7 / max(len(lunch_pool), 1)))
    cap_d = max(cfg["max_repeats_dinner"], math.ceil(7 / max(len(dinner_pool), 1)))
    return {"cap_l": cap_l, "cap_d": cap_d}


# ----------------------------------------------------------------------------- shopping list (pure Python)
def build_matchers(grocery, cfg):
    """One regex per grocery item, also matching its aliases (palak -> Spinach)."""
    alias_for = defaultdict(set)
    for a, target in cfg["aliases"].items():
        alias_for[norm(target)].add(norm(a))
    matchers, seen = [], set()
    for g in grocery:
        k = item_key(g["item"], cfg["aliases"])
        if k in seen:
            continue
        seen.add(k)
        names = sorted({norm(g["item"])} | alias_for[norm(g["item"])], key=len, reverse=True)
        rx = re.compile("|".join(rf"\b{re.escape(n)}s?\b" for n in names))
        matchers.append({"item": g["item"], "category": g["category"], "key": k, "rx": rx})
    return matchers


def extra_ingredients(text, matchers, ignore) -> list[str]:
    out = []
    for tok in re.split(r"[,;]", str(text or "")):
        t = re.sub(r"\(.*?\)", "", norm(tok))
        t = re.sub(r"^[\d½¼¾/.\-– ]+(kg|g|ml|l|cups?|tsp|tbsp)?\b", "", t).strip()
        if len(t) < 3 or any(m["rx"].search(t) for m in matchers) or any(i in t for i in ignore):
            continue
        out.append(t)
    return out


def wants_rice(name: str, cfg: dict) -> bool:
    n = norm(name)
    return bool(cfg["serve_rice"]) and "rice" not in n and any(k in n for k in cfg["rice_with"])


def build_shopping(days, dishes, matchers, pantry, cfg, snacks_text=None):
    by_name = {norm(d["name"]): d for d in dishes}
    needs: dict[str, dict] = {}
    no_ingredients: set[str] = set()

    def add(display, category, key, label, times=1, note=""):
        e = needs.setdefault(key, {"item": display, "category": category, "uses": [], "times": 0, "notes": set()})
        e["times"] += times
        if label and label not in e["uses"]:
            e["uses"].append(label)
        if note:
            e["notes"].add(note)

    for k in (item_key(s, cfg["aliases"]) for s in cfg["weekly_staples"]):
        m = next((m for m in matchers if m["key"] == k), None)
        disp = m["item"] if m else k.title()
        add(disp, m["category"] if m else "Not in Grocery sheet", k, "Breakfast", 7)

    for day in days:
        for slot in ("lunch", "dinner"):
            dish = by_name.get(norm(day[slot]))
            if not dish:
                continue
            text = norm(dish["name"] + " " + dish["ingredients"])
            for m in matchers:
                if m["rx"].search(text):
                    add(m["item"], m["category"], m["key"], dish["name"])
            for extra in extra_ingredients(dish["ingredients"], matchers, cfg["ignore_ingredients"]):
                add(extra.title(), "Not in Grocery sheet", item_key(extra, cfg["aliases"]), dish["name"],
                    note="Not in your Grocery sheet - add it there if you buy it")
            if not dish["ingredients"]:
                no_ingredients.add(dish["name"])
            if wants_rice(dish["name"], cfg):
                add("Rice", "Pulses & grains", "rice", dish["name"])
        snack = norm(day.get("snack"))
        matched = [m for m in matchers if m["rx"].search(snack)]
        for m in matched:
            add(m["item"], m["category"], m["key"], "Snack")
        if snack and not matched and "fruit" not in snack:
            add(day["snack"], "Not in Grocery sheet", item_key(day["snack"], cfg["aliases"]), "Snack",
                note="Snack not in your Grocery sheet")

    rows = []
    for key, e in needs.items():
        if key in pantry:
            status, stock = "In stock", pantry[key]["qty"]
        elif e["category"] == "Not in Grocery sheet":
            status, stock = "Check / optional", ""
        else:
            status, stock = "Buy", ""
        note = "; ".join(sorted(e["notes"]))
        if status == "In stock":
            note = ("Verify there is enough for the week. " + note).strip()
        rows.append([e["item"], e["category"], status, ", ".join(e["uses"]), e["times"], stock, note, ""])
    order = {"Buy": 0, "Check / optional": 1, "In stock": 2}
    rows.sort(key=lambda r: (order[r[2]], r[1], r[0]))
    return rows, sorted(no_ingredients)


# ----------------------------------------------------------------------------- writing results
HEADER_FILL = PatternFill("solid", fgColor="365F91")
FILLS = {"Buy": "FCE4D6", "Check / optional": "EDEDED", "In stock": "E2EFDA"}


def style_sheet(ws, filt=True):
    ws.freeze_panes = "A2"
    if filt:
        ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, ws.max_column + 1):
        width = 12
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row, col).value
            if v is not None:
                width = max(width, min(len(str(v)), 55))
                ws.cell(row, col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(col)].width = min(width + 2, 50)


def write_outputs(wb, days, shop_rows, plan, dishes, cfg, monday, notes):
    for name in OUTPUT_SHEETS:
        if name in wb.sheetnames:
            del wb[name]
    status = {d["name"]: d["status"] for d in dishes}

    ws = wb.create_sheet("Meal Plan")
    ws.append(["Day", "Date", "Breakfast", "Lunch", "Dinner", "Snack", "Why this plan"])
    for i, d in enumerate(days):
        learn = status.get(d["dinner"]) != "known"
        lunch = d["lunch"] + (" + rice" if wants_rice(d["lunch"], cfg) else "")
        dinner = d["dinner"] + (" + rice" if wants_rice(d["dinner"], cfg) else "") + (" (learn)" if learn else "")
        ws.append([d["day"], (monday + timedelta(days=i)).strftime("%d %b"), cfg["fixed_breakfast"],
                   lunch, dinner, d["snack"], d["reason"]])
        if learn:
            ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor="FFF2CC")
    style_sheet(ws, filt=False)

    ws = wb.create_sheet("Shopping List")
    ws.append(["Item", "Category", "Status", "Needed for", "Meals this week", "In pantry", "Notes", "Bought?"])
    for r in shop_rows:
        ws.append(r)
        ws.cell(ws.max_row, 3).fill = PatternFill("solid", fgColor=FILLS[r[2]])
    style_sheet(ws)

    # ws = wb.create_sheet("Prep Notes")
    # ws.append([f"Week of {monday.strftime('%d %b %Y')}"])
    # for n in (plan.get("prep_notes") or []) + notes:
    #     ws.append([str(n)])
    # style_sheet(ws, filt=False)
    # ws.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Meal Prep")
    for line in meal_prep_text(days, dishes, shop_rows, cfg, monday).splitlines():
        ws.append([line])
    style_sheet(ws, filt=False)
    ws.column_dimensions["A"].width = 110

# ----------------------------------------------------------------------------- engine (used by CLI and UI)
def load_context(wb, cfg) -> SimpleNamespace:
    grocery = read_grocery_candidates(get_sheet(wb, "Grocery"))
    pantry = read_pantry(ensure_pantry_sheet(wb), cfg["aliases"])
    dishes = read_dishes(get_sheet(wb, "What can i make"))
    snacks = read_snacks(get_sheet(wb, "snacks"))
    nutrition_priorities = read_nutrition_priorities(get_sheet(wb, "required", False))
    lunch_pool, dinner_pool, excluded = classify_dishes(dishes, cfg)
    if not lunch_pool or not dinner_pool:
        raise ValueError("No usable dishes: mark some dishes 'Known' (lunch) and give them a meal type.")
    return SimpleNamespace(grocery=grocery, pantry=pantry, dishes=dishes, snacks=snacks,
                           nutrition_priorities=nutrition_priorities,
                           lunch_pool=lunch_pool, dinner_pool=dinner_pool, excluded=excluded,
                           lim=limits(cfg, lunch_pool, dinner_pool), matchers=build_matchers(grocery, cfg))


def generate(ctx, cfg, locked=None, rules=None):
    """Choose a varied week. Active user rules are enforced before scoring."""
    rng = random.Random()
    locked = locked or {}
    rules = [r for r in (rules or []) if r.get("active", True)]
    all_dishes = {norm(d["name"]): d for d in ctx.dishes}
    day_index = {name.lower(): i for i, name in enumerate(DAYS)}
    fixed = {}
    avoided_dishes = {norm(r.get("dish") or r.get("value")) for r in rules if r.get("rule_type") == "avoid_dish"}
    avoided_terms = [norm(r.get("value")) for r in rules if r.get("rule_type") == "avoid_ingredient" and r.get("value")]
    preferred_terms = [norm(r.get("value")) for r in rules if r.get("rule_type") == "prefer_ingredient" and r.get("value")]
    for rule in rules:
        if rule.get("rule_type") != "fixed_meal":
            continue
        idx = day_index.get(norm(rule.get("day")))
        slot = norm(rule.get("meal"))
        dish = str(rule.get("dish") or rule.get("value") or "").strip()
        if idx is None or slot not in {"lunch", "dinner"} or not dish:
            continue
        if norm(dish) not in all_dishes:
            raise ValueError(f"Your fixed-meal rule refers to '{dish}', which is not in the workbook.")
        pool = ctx.lunch_pool if slot == "lunch" else ctx.dinner_pool
        if norm(dish) not in {norm(d["name"]) for d in pool}:
            raise ValueError(f"'{dish}' is not currently eligible for {slot} under your workbook settings.")
        fixed[(idx, slot)] = dish
    traits = {}
    for key, dish in all_dishes.items():
        text = norm(dish["name"] + " " + dish["ingredients"])
        traits[key] = {priority["name"] for priority in ctx.nutrition_priorities
                       if any(food in text for food in priority["foods"])}

    def dish_score(name):
        dish = all_dishes.get(norm(name), {})
        ingredients = norm(dish.get("name", "") + " " + dish.get("ingredients", ""))
        preference_score = sum(2 for term in preferred_terms if term in ingredients)
        return (3 if dish.get("status") == "known" else 0) + {
            "Easy": 2, "Medium": 1, "Difficult": -3,
        }.get(dish.get("ease"), 0) + preference_score

    def added_traits(names, counts):
        score = 0
        updated = counts.copy()
        targets = {priority["name"]: priority["target"] for priority in ctx.nutrition_priorities}
        for name in names:
            for nutrient in traits.get(norm(name), set()):
                if updated[nutrient] < targets[nutrient]:
                    score += targets[nutrient] - updated[nutrient]
                    updated[nutrient] += 1
        return score, updated

    def search(day_index, selected, use_lunch, use_dinner, nutrient_counts):
        if day_index == 7:
            return selected
        pin = locked.get(day_index, {})
        lunches = [fixed[(day_index, "lunch")]] if (day_index, "lunch") in fixed else ([pin["lunch"]] if pin.get("lunch") else [d["name"] for d in ctx.lunch_pool])
        dinners = [fixed[(day_index, "dinner")]] if (day_index, "dinner") in fixed else ([pin["dinner"]] if pin.get("dinner") else [d["name"] for d in ctx.dinner_pool])
        def allowed(name):
            dish = all_dishes.get(norm(name), {})
            text = norm(dish.get("name", "") + " " + dish.get("ingredients", ""))
            return norm(name) not in avoided_dishes and not any(term and term in text for term in avoided_terms)
        lunches = [name for name in lunches if allowed(name)]
        dinners = [name for name in dinners if allowed(name)]
        options = [(lunch, dinner) for lunch in lunches for dinner in dinners
                   if lunch != dinner and use_lunch[lunch] < ctx.lim["cap_l"]
                   and use_dinner[dinner] < ctx.lim["cap_d"]
                   and (not selected or lunch not in selected[-1][:2] and dinner not in selected[-1][:2])]
        rng.shuffle(options)
        options.sort(key=lambda pair: dish_score(pair[0]) + dish_score(pair[1])
                     + added_traits(pair, nutrient_counts)[0], reverse=True)
        for lunch, dinner in options:
            _, updated_nutrients = added_traits((lunch, dinner), nutrient_counts)
            selected.append((lunch, dinner))
            use_lunch[lunch] += 1
            use_dinner[dinner] += 1
            result = search(day_index + 1, selected, use_lunch, use_dinner, updated_nutrients)
            if result is not None:
                return result
            selected.pop()
            use_lunch[lunch] -= 1
            use_dinner[dinner] -= 1
        return None

    fixed_breakfast = norm(cfg["fixed_breakfast"])
    nutrition_counts = Counter()
    for priority in ctx.nutrition_priorities:
        if any(food in fixed_breakfast for food in priority["foods"]):
            nutrition_counts[priority["name"]] = len(DAYS)
    selected = search(0, [], Counter(), Counter(), nutrition_counts)
    if selected is None:
        raise ValueError("Could not build a week with those locked meals and repeat limits. Unlock a day or add eligible dishes.")

    days = []
    for i, (lunch, dinner) in enumerate(selected):
        snack = locked.get(i, {}).get("snack") or (rng.choice(ctx.snacks) if ctx.snacks else "")
        days.append({"day": DAYS[i], "lunch": lunch, "dinner": dinner, "snack": snack,
                     "reason": "Selected using your workbook's nutrient priorities."})
    return {"prep_notes": [], "suggested_new_dishes": []}, days, []


def analyse(days, ctx, cfg):
    """Shopping list + warnings for the current plan (cheap, no model call; safe to re-run after edits)."""
    rows, no_ing = build_shopping(days, ctx.dishes, ctx.matchers, ctx.pantry, cfg)
    warnings = []
    if len(ctx.lunch_pool) < 4:
        warnings.append(f"Only {len(ctx.lunch_pool)} Known lunch dishes, so lunches repeat. Mark more dishes as Known "
                        "(and set their meal type to Lunch) to get more variety.")
    if no_ing:
        warnings.append("No ingredient list in the workbook for: " + ", ".join(no_ing) +
                        ". Their shopping items are guessed from the dish name only - fill in the Ingredients column.")
    return rows, warnings


def meal_prep_text(days, dishes, shopping_rows, cfg, monday):
    by_name = {norm(dish["name"]): dish for dish in dishes}
    buy = [row for row in shopping_rows if row[2] == "Buy"]
    check = [row for row in shopping_rows if row[2] == "Check / optional"]
    missing_ingredients = set()
    for day in days:
        for slot in ("lunch", "dinner"):
            dish = by_name.get(norm(day[slot]))
            if dish and not dish["ingredients"]:
                missing_ingredients.add(dish["name"])

    lines = [f"## Meal prep and pantry check - week of {monday.strftime('%d %b %Y')}", "",
             "## Buy - not found in pantry"]
    lines.extend([f"- {row[0]} (needed for: {row[3]})" for row in buy] or ["- Nothing identified."])
    lines += ["", "## Check manually - not in Grocery sheet"]
    lines.extend([f"- {row[0]} (needed for: {row[3]})" for row in check] or ["- Nothing identified."])
    lines += ["", "## Daily meal prep"]
    for index, day in enumerate(days):
        day_date = (monday + timedelta(days=index)).strftime('%a %d %b')
        lines += [f"### {day_date}", f"- Breakfast: {cfg['fixed_breakfast']}"]
        for slot in ("lunch", "dinner"):
            name = day[slot]
            dish = by_name.get(norm(name))
            meal = name + (" + rice" if wants_rice(name, cfg) else "")
            lines.append(f"- {slot.title()}: {meal}")
            if dish and dish["ingredients"]:
                lines.append(f"  - Ingredients: {dish['ingredients']}")
            else:
                lines.append("  - Ingredients not listed in workbook.")
        lines.append(f"- Snack: {day.get('snack', '')}")
        lines.append("")
    lines.append("## Workbook ingredient gaps")
    lines.extend([f"- {name}: add ingredients in 'What can i make' for a complete pantry check."
                  for name in sorted(missing_ingredients)] or ["- No missing main-meal ingredient lists."])
    return "\n".join(lines).strip()


def plan_monday(today: date | None = None) -> date:
    today = today or date.today()
    return today + timedelta(days=(0 - today.weekday()) % 7)  # Sunday run -> planning for tomorrow


def write_pantry(wb, rows: list[dict]):
    """Replace the Pantry rows with `rows` = [{item, qty, unit}] (header untouched)."""
    ws = ensure_pantry_sheet(wb)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for r in rows:
        if str(r.get("item") or "").strip():
            ws.append([str(r["item"]).strip(), r.get("qty"), r.get("unit") or ""])


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Build a weekly meal plan and shopping list from an Excel workbook.")
    ap.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    ap.add_argument("--inplace", action="store_true", help="Write the plan into the workbook itself")
    args = ap.parse_args()

    path = Path(args.workbook).expanduser().resolve()
    if not path.exists():
        raise SystemExit(f"Workbook not found: {path}")

    wb = load_workbook(path)
    cfg = load_config(wb)
    try:
        ctx = load_context(wb, cfg)
    except ValueError as exc:
        raise SystemExit(str(exc))
    print(f"Lunch options (Known dishes): {len(ctx.lunch_pool)} | Dinner options: {len(ctx.dinner_pool)} | Pantry items: {len(ctx.pantry)}")
    for name, why in ctx.excluded:
        print(f"  skipped: {name} ({why})")

    plan, days, repair_notes = generate(ctx, cfg)
    shop_rows, warnings = analyse(days, ctx, cfg)
    warnings = repair_notes + warnings
    write_outputs(wb, days, shop_rows, plan, ctx.dishes, cfg, plan_monday(), warnings)

    if args.inplace:
        out = path
    else:
        out = path.with_name(f"{path.stem}_planned_{date.today().isoformat()}{path.suffix}")
    wb.save(out)
    print(f"\nDone -> {out}")
    for w in warnings:
        print(f"  note: {w}")
    print("Review 'Meal Plan' and 'Shopping List' before shopping.")


if __name__ == "__main__":
    main()