"""
Weekly meal planner: Excel + local Ollama.

Division of labour (this is what makes a 7B local model reliable):
  * Python  -> reads the workbook, applies your hard rules (no roti, known dishes
               for lunch, learning dishes for dinner, meal types), validates the
               model's answer, retries with feedback, and builds the shopping list.
  * Ollama  -> only chooses lunch / dinner / snack per day and explains why.
"""
from __future__ import annotations

import argparse
import difflib
import json
import math
import re
import shutil
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
CONFIG_PATH = ROOT / "config.json"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
OUTPUT_SHEETS = ("Meal Plan", "Shopping List", "Prep Notes", "Suggestions")

DEFAULT_CONFIG: dict[str, Any] = {
    "model": "qwen2.5:7b",
    "temperature": 0.3,
    "num_ctx": 8192,              # Ollama's default context is small; the prompt needs more
    "max_retries": 2,             # extra attempts when the model breaks a rule
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
    "min_learning_dinners": 4,
    "greens": ["Spinach", "Methi"],
    "diet": "Mostly vegetarian, occasional meat; eggs and dairy are okay.",
    "health_context": "Prioritise folate-rich greens/legumes and protein. Low folate, markedly low vitamin D "
                      "and low-normal B12 were reported. Food supports but does not replace clinician-directed treatment.",
    "activity": "Mostly sitting, little exercise.",
    # words in an ingredient list that you always have / never need to buy
    "ignore_ingredients": ["salt", "oil", "powder", "masala", "jeera", "pepper", "chilli", "bay leaves",
                           "cinamaon", "cinnamon", "nutmeg", "turmeric", "corrainder", "corriander", "coriander",
                           "mirchi", "pasre", "lemon juice", "water", "dal"],
    # spelling/synonym fixes: workbook word -> Grocery sheet item
    "aliases": {"palak": "Spinach", "rajma": "Rajama", "soya chunk": "Soy chunks",
                "soya chunks": "Soy chunks", "soy chunk": "Soy chunks", "almonds": "Almond"},
}


# ----------------------------------------------------------------------------- helpers
def load_config() -> dict:
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=2), encoding="utf-8")
    cfg = dict(DEFAULT_CONFIG)
    cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
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


def read_reference(ws) -> list[str]:
    """Optional 'required' / 'notes' sheets: passed to the model as nutrition context."""
    if ws is None:
        return []
    rows = []
    for r in ws.iter_rows(values_only=True):
        cells = [str(c).strip() for c in r if c is not None and str(c).strip()]
        if cells:
            rows.append(" | ".join(cells))
    return rows


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
    learn = sum(1 for d in dinner_pool if d["status"] != "known")
    min_learn = min(cfg["min_learning_dinners"], learn * cap_d, 7)
    return {"cap_l": cap_l, "cap_d": cap_d, "min_learn": min_learn}


def brief(d: dict) -> dict:
    out = {"name": d["name"], "ease": d["ease"], "status": d["status"]}
    if d["serving"]:
        out["serving"] = d["serving"]
    if d["ingredients"]:
        out["ingredients"] = d["ingredients"][:110]
    return out


def build_prompt(cfg, lunch_pool, dinner_pool, snacks, grocery, pantry, reference, lim, locked=None) -> str:
    shape = {
        "days": [{"day": "Monday", "lunch": "<exact name from LUNCH_OPTIONS>",
                  "dinner": "<exact name from DINNER_OPTIONS>", "snack": "<one item from SNACKS>",
                  "reason": "<one short sentence>"}, "...exactly 7 entries, Monday..Sunday"],
        "prep_notes": ["<short practical tip>"],
        "suggested_new_dishes": [{"name": "<dish>", "meal": "Dinner", "why": "<short>",
                                  "main_ingredients": ["<items from the Grocery list>"]}],
    }
    context = {
        "LUNCH_OPTIONS": [brief(d) for d in lunch_pool],
        "DINNER_OPTIONS": [brief(d) for d in dinner_pool],
        "SNACKS": snacks,
        "GROCERY_LIST_ITEMS": sorted({g["item"] for g in grocery}),
        "ALREADY_AT_HOME": [p["item"] for p in pantry.values()],
        "NUTRITION_REFERENCE": reference,
    }
    locked_txt = ""
    if locked:
        parts = [f"{DAYS[i]}: " + ", ".join(f"{k}={v}" for k, v in v_.items() if v) for i, v_ in sorted(locked.items())]
        locked_txt = ("\nThe user has already fixed these slots - repeat them exactly and plan the rest around them "
                      "(avoid repeating them elsewhere):\n- " + "\n- ".join(parts) + "\n")
    return f"""Plan lunch, dinner and a snack for 7 days (Monday-Sunday). Reply with ONLY JSON in this shape:
{json.dumps(shape, ensure_ascii=False)}

About the user: {cfg['diet']} {cfg['activity']} {cfg['health_context']}
Breakfast is fixed ({cfg['fixed_breakfast']}) - do not plan it. Rice is cooked as a side automatically.
{locked_txt}
Hard rules:
- "lunch" must be copied EXACTLY from LUNCH_OPTIONS; "dinner" EXACTLY from DINNER_OPTIONS.
- Lunch and dinner on the same day must be different dishes.
- No lunch dish more than {lim['cap_l']} times and no dinner dish more than {lim['cap_d']} time(s) in the week.
- At least {lim['min_learn']} dinners must have status "learning", "not tried" or "unknown" (the user wants to learn new dishes at dinner).
Soft rules:
- Vary the week. Include leafy greens (spinach, methi) on several days, pulses/legumes, and a protein source in every main meal (paneer, soy chunks, eggs, dal, rajma, curd, occasional chicken).
- Prefer snacks that use items already in the Grocery list. No calorie-restriction or weight-loss framing; no supplement doses.
- "suggested_new_dishes": 0-3 simple dinner ideas NOT in the option lists, built mostly from GROCERY_LIST_ITEMS.
- "prep_notes": 3-5 short tips (e.g. soak rajma overnight, batch-wash greens).

DATA:
{json.dumps(context, ensure_ascii=False, separators=(",", ":"))}
"""


# ----------------------------------------------------------------------------- Ollama + validation
def extract_json(text: str) -> dict:
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    raise ValueError("The reply was not valid JSON.")


def chat(messages, cfg) -> str:
    try:
        import ollama
    except ImportError as exc:
        raise SystemExit("The 'ollama' package is missing. Run: pip install -r requirements.txt") from exc
    try:
        resp = ollama.chat(model=cfg["model"], messages=messages, format="json",
                           options={"temperature": float(cfg["temperature"]), "num_ctx": int(cfg["num_ctx"])})
    except Exception as exc:  # connection refused, model not pulled, ...
        raise SystemExit(f"Could not talk to Ollama ({exc}).\n"
                         f"Is the Ollama app running, and have you run `ollama pull {cfg['model']}`?") from exc
    return resp["message"]["content"]


def lookup(name, mapping: dict):
    """Exact match first, then a close spelling match ('panner bhurji' -> 'paneer bhurji')."""
    n = norm(name)
    if n in mapping:
        return mapping[n]
    close = difflib.get_close_matches(n, list(mapping), n=1, cutoff=0.8)
    return mapping[close[0]] if close else None


def check_plan(plan, lunch_pool, dinner_pool, lim):
    """Returns (errors, cleaned_days or None). Cleaned days use canonical dish names."""
    days = plan.get("days") if isinstance(plan, dict) else None
    if not isinstance(days, list) or len(days) != 7 or not all(isinstance(d, dict) for d in days):
        return ["'days' must be a list of exactly 7 objects (Monday..Sunday)."], None
    l_map = {norm(d["name"]): d["name"] for d in lunch_pool}
    d_map = {norm(d["name"]): d["name"] for d in dinner_pool}
    learn = {d["name"] for d in dinner_pool if d["status"] != "known"}
    errors, cleaned = [], []
    ul, ud = Counter(), Counter()
    for i, day in enumerate(days):
        lunch, dinner = lookup(day.get("lunch"), l_map), lookup(day.get("dinner"), d_map)
        if not lunch:
            errors.append(f"{DAYS[i]}: lunch {day.get('lunch')!r} is not in LUNCH_OPTIONS.")
        if not dinner:
            errors.append(f"{DAYS[i]}: dinner {day.get('dinner')!r} is not in DINNER_OPTIONS.")
        if lunch and lunch == dinner:
            errors.append(f"{DAYS[i]}: lunch and dinner are the same dish.")
        ul[lunch] += 1
        ud[dinner] += 1
        cleaned.append({"day": DAYS[i], "lunch": lunch or "", "dinner": dinner or "",
                        "snack": str(day.get("snack", "")), "reason": str(day.get("reason", ""))})
    errors += [f"Lunch '{n}' used {c} times (max {lim['cap_l']})." for n, c in ul.items() if n and c > lim["cap_l"]]
    errors += [f"Dinner '{n}' used {c} times (max {lim['cap_d']})." for n, c in ud.items() if n and c > lim["cap_d"]]
    n_learn = sum(1 for d in cleaned if d["dinner"] in learn)
    if n_learn < lim["min_learn"]:
        errors.append(f"Only {n_learn} dinners are learning/new dishes; need at least {lim['min_learn']}.")
    return errors, cleaned


def repair(cleaned, lunch_pool, dinner_pool, lim, snacks):
    """Last resort: swap in the least-used valid dish for any slot that breaks a rule."""
    notes = []
    cleaned = cleaned or [{"day": d, "lunch": "", "dinner": "", "snack": "", "reason": ""} for d in DAYS]
    for slot, pool, cap in (("lunch", lunch_pool, lim["cap_l"]), ("dinner", dinner_pool, lim["cap_d"])):
        names = [d["name"] for d in pool]
        learn = {d["name"] for d in pool if d["status"] != "known"}
        use = Counter()
        for day in cleaned:
            cur = day[slot]
            bad = cur not in names or use[cur] >= cap or (slot == "dinner" and cur == day["lunch"])
            if bad:
                options = [n for n in names if not (slot == "dinner" and n == day["lunch"])]
                options.sort(key=lambda n: (use[n], n not in learn if slot == "dinner" else 0))
                day[slot] = options[0]
                notes.append(f"{day['day']} {slot} was auto-corrected to '{options[0]}'.")
            use[day[slot]] += 1
    for i, day in enumerate(cleaned):
        if not day["snack"] and snacks:
            day["snack"] = snacks[i % len(snacks)]
    return cleaned, notes


def get_plan(prompt, cfg, lunch_pool, dinner_pool, snacks, lim, log=print):
    messages = [
        {"role": "system", "content": "You are a careful meal-planning assistant. Follow every rule. Output valid JSON only."},
        {"role": "user", "content": prompt},
    ]
    plan, cleaned, errors = {}, None, []
    for attempt in range(1, int(cfg["max_retries"]) + 2):
        log(f"Asking {cfg['model']} (attempt {attempt}) ...")
        raw = chat(messages, cfg)
        try:
            plan = extract_json(raw)
            errors, cleaned = check_plan(plan, lunch_pool, dinner_pool, lim)
        except ValueError as exc:
            errors, cleaned = [str(exc)], None
        if not errors:
            return plan, cleaned, []
        log("  Rule problems found:\n   - " + "\n   - ".join(errors[:8]))
        messages += [{"role": "assistant", "content": raw},
                     {"role": "user", "content": "Fix these problems and return the complete corrected JSON only:\n- "
                      + "\n- ".join(errors)}]
    cleaned, notes = repair(cleaned, lunch_pool, dinner_pool, lim, snacks)
    return plan if isinstance(plan, dict) else {}, cleaned, notes


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
    greens = {item_key(g, cfg["aliases"]) for g in cfg["greens"]}
    greens_days = 0
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
        day_has_greens = False
        for slot in ("lunch", "dinner"):
            dish = by_name.get(norm(day[slot]))
            if not dish:
                continue
            text = norm(dish["name"] + " " + dish["ingredients"])
            for m in matchers:
                if m["rx"].search(text):
                    add(m["item"], m["category"], m["key"], dish["name"])
                    day_has_greens |= m["key"] in greens
            for extra in extra_ingredients(dish["ingredients"], matchers, cfg["ignore_ingredients"]):
                add(extra.title(), "Not in Grocery sheet", item_key(extra, cfg["aliases"]), dish["name"],
                    note="Not in your Grocery sheet - add it there if you buy it")
            if not dish["ingredients"]:
                no_ingredients.add(dish["name"])
            if wants_rice(dish["name"], cfg):
                add("Rice", "Pulses & grains", "rice", dish["name"])
        greens_days += day_has_greens
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
    return rows, greens_days, sorted(no_ingredients)


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

    ws = wb.create_sheet("Prep Notes")
    ws.append([f"Week of {monday.strftime('%d %b %Y')}"])
    for n in (plan.get("prep_notes") or []) + notes:
        ws.append([str(n)])
    style_sheet(ws, filt=False)
    ws.column_dimensions["A"].width = 110

    ideas = plan.get("suggested_new_dishes") or []
    if ideas:
        ws = wb.create_sheet("Suggestions")
        ws.append(["Dish idea", "Meal", "Why", "Main ingredients", "Copy into 'What can i make'?"])
        for s in ideas:
            if isinstance(s, dict):
                mi = s.get("main_ingredients")
                ws.append([s.get("name", ""), s.get("meal", "Dinner"), s.get("why", ""),
                           ", ".join(mi) if isinstance(mi, list) else str(mi or ""), ""])
        style_sheet(ws)


# ----------------------------------------------------------------------------- engine (used by CLI and UI)
def load_context(wb, cfg) -> SimpleNamespace:
    grocery = read_grocery_candidates(get_sheet(wb, "Grocery"))
    pantry = read_pantry(ensure_pantry_sheet(wb), cfg["aliases"])
    dishes = read_dishes(get_sheet(wb, "What can i make"))
    snacks = read_snacks(get_sheet(wb, "snacks"))
    reference = read_reference(get_sheet(wb, "required", False))
    lunch_pool, dinner_pool, excluded = classify_dishes(dishes, cfg)
    if not lunch_pool or not dinner_pool:
        raise ValueError("No usable dishes: mark some dishes 'Known' (lunch) and give them a meal type.")
    return SimpleNamespace(grocery=grocery, pantry=pantry, dishes=dishes, snacks=snacks, reference=reference,
                           lunch_pool=lunch_pool, dinner_pool=dinner_pool, excluded=excluded,
                           lim=limits(cfg, lunch_pool, dinner_pool), matchers=build_matchers(grocery, cfg))


def generate(ctx, cfg, locked=None, log=print):
    """Ask the model for a week. `locked` = {day_index: {"lunch":..,"dinner":..,"snack":..}} kept as-is."""
    prompt = build_prompt(cfg, ctx.lunch_pool, ctx.dinner_pool, ctx.snacks, ctx.grocery, ctx.pantry,
                          ctx.reference, ctx.lim, locked)
    plan, days, notes = get_plan(prompt, cfg, ctx.lunch_pool, ctx.dinner_pool, ctx.snacks, ctx.lim, log)
    for i, vals in (locked or {}).items():
        days[i].update({k: v for k, v in vals.items() if v})
    return plan, days, notes


def analyse(days, ctx, cfg):
    """Shopping list + warnings for the current plan (cheap, no model call; safe to re-run after edits)."""
    rows, greens_days, no_ing = build_shopping(days, ctx.dishes, ctx.matchers, ctx.pantry, cfg)
    warnings = []
    if len(ctx.lunch_pool) < 4:
        warnings.append(f"Only {len(ctx.lunch_pool)} Known lunch dishes, so lunches repeat. Mark more dishes as Known "
                        "(and set their meal type to Lunch) to get more variety.")
    if greens_days < 3:
        warnings.append(f"Leafy greens appear on only {greens_days} day(s). Folate-rich greens matter for you - "
                        "mark more spinach/methi dishes as Known or add meal types.")
    if no_ing:
        warnings.append("No ingredient list in the workbook for: " + ", ".join(no_ing) +
                        ". Their shopping items are guessed from the dish name only - fill in the Ingredients column.")
    return rows, warnings


def plan_monday(today: date | None = None) -> date:
    today = today or date.today()
    return today + timedelta(days=(0 - today.weekday()) % 7)  # Sunday run -> planning for tomorrow


def backup_once(path: Path) -> Path:
    backup = path.with_name(f"{path.stem}_backup_{date.today().isoformat()}{path.suffix}")
    if not backup.exists():
        shutil.copy2(path, backup)
    return backup


# --- workbook editing helpers (the UI uses these; each takes an open openpyxl workbook)
DISH_FIELDS = {"name": ["Dish", "Recipe", "Name"], "link": ["Link", "URL"], "ingredients": ["Ingredients"],
               "ease": ["Ease (E/M/D)", "Ease", "Difficulty"], "serving": ["Serving size (1 person)", "Serving size", "Portion"],
               "meal": ["B / L / D", "Meal", "Meal type"], "status": ["Cooking status", "Familiarity"]}


def dish_columns(ws) -> dict[str, int]:
    headers = [c.value for c in ws[1]]
    return {k: i + 1 for k, v in DISH_FIELDS.items() if (i := find_col(headers, v)) is not None}


def update_dish_cells(wb, changes: list[tuple[int, str, Any]]):
    """changes = [(excel_row, field, value)] - only the cells you pass are touched."""
    ws = get_sheet(wb, "What can i make")
    cols = dish_columns(ws)
    for row, field, value in changes:
        if field in cols:
            ws.cell(row, cols[field]).value = value if value != "" else None


def append_dish(wb, dish: dict):
    ws = get_sheet(wb, "What can i make")
    cols = dish_columns(ws)
    last = max((r for r in range(2, ws.max_row + 1) if ws.cell(r, cols["name"]).value), default=1)
    for field, value in dish.items():
        if field in cols and value not in (None, ""):
            ws.cell(last + 1, cols[field]).value = value


def write_pantry(wb, rows: list[dict]):
    """Replace the Pantry rows with `rows` = [{item, qty, unit}] (header untouched)."""
    ws = ensure_pantry_sheet(wb)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    for r in rows:
        if str(r.get("item") or "").strip():
            ws.append([str(r["item"]).strip(), r.get("qty"), r.get("unit") or ""])


def meals_text(meals: set) -> str:
    order = [("b", "Breakfast"), ("l", "Lunch"), ("d", "Dinner")]
    return " / ".join(name for k, name in order if k in meals)


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Weekly meal plan + shopping list from Grocery.xlsx via local Ollama.")
    ap.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    ap.add_argument("--model", default=None, help="Override the Ollama model, e.g. llama3.1:8b")
    ap.add_argument("--inplace", action="store_true",
                    help="Write the plan into the workbook itself (a dated backup is made first)")
    ap.add_argument("--show-prompt", action="store_true", help="Print the prompt and exit (no model call)")
    args = ap.parse_args()

    cfg = load_config()
    if args.model:
        cfg["model"] = args.model
    path = Path(args.workbook).expanduser().resolve()
    if not path.exists():
        raise SystemExit(f"Workbook not found: {path}")

    wb = load_workbook(path)
    try:
        ctx = load_context(wb, cfg)
    except ValueError as exc:
        raise SystemExit(str(exc))
    print(f"Lunch options (Known dishes): {len(ctx.lunch_pool)} | Dinner options: {len(ctx.dinner_pool)} | Pantry items: {len(ctx.pantry)}")
    for name, why in ctx.excluded:
        print(f"  skipped: {name} ({why})")

    if args.show_prompt:
        print(build_prompt(cfg, ctx.lunch_pool, ctx.dinner_pool, ctx.snacks, ctx.grocery, ctx.pantry, ctx.reference, ctx.lim))
        return

    plan, days, repair_notes = generate(ctx, cfg)
    shop_rows, warnings = analyse(days, ctx, cfg)
    warnings = repair_notes + warnings
    write_outputs(wb, days, shop_rows, plan, ctx.dishes, cfg, plan_monday(), warnings)

    if args.inplace:
        out = path
        print(f"Backup saved: {backup_once(path).name}")
    else:
        out = path.with_name(f"{path.stem}_planned_{date.today().isoformat()}{path.suffix}")
    wb.save(out)
    print(f"\nDone -> {out}")
    for w in warnings:
        print(f"  note: {w}")
    print("Review 'Meal Plan' and 'Shopping List' before shopping.")


if __name__ == "__main__":
    main()