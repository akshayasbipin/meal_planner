"""Interactive meal planner UI.  Run with:  streamlit run app.py"""
from __future__ import annotations

import json
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import meal_planner as mp
import ui_logic as ui

st.set_page_config(page_title="Weekly Meal Planner", page_icon="🥗", layout="wide")
ss = st.session_state
PLANS_DIR = mp.ROOT / "plans"


# ------------------------------------------------------------------ state / io helpers
def init_state():
    ss.setdefault("cfg", mp.load_config())
    ss.setdefault("days", None)
    ss.setdefault("plan", {})
    ss.setdefault("notes", [])
    ss.setdefault("monday", mp.plan_monday())
    ss.setdefault("week_start", ss.monday)
    ss.setdefault("loaded_for", None)
    ss.setdefault("log", [])
    ss.setdefault("snap", "")


def set_plan(days, plan, notes, monday):
    ss.days, ss.plan, ss.notes, ss.monday = days, plan, notes, monday
    ss.loaded_for = monday
    ss.snap = json.dumps(days)
    for i, d in enumerate(days):  # push into the widgets
        for slot in ("lunch", "dinner", "snack"):
            ss[f"{slot}_{i}"] = d[slot]


def plan_file(monday: date) -> Path:
    return PLANS_DIR / f"plan_{monday.isoformat()}.json"


def save_json():
    PLANS_DIR.mkdir(exist_ok=True)
    plan_file(ss.monday).write_text(json.dumps(
        {"monday": ss.monday.isoformat(), "days": ss.days, "plan": ss.plan, "notes": ss.notes}, indent=2), encoding="utf-8")


def load_json(path: Path):
    data = json.loads(path.read_text(encoding="utf-8"))
    set_plan(data["days"], data.get("plan", {}), data.get("notes", []), date.fromisoformat(data["monday"]))
    ss.week_start = ss.monday


def change_week():
    selected = ss.week_start
    monday = selected - timedelta(days=selected.weekday())
    if monday == ss.monday:
        ss.week_start = monday
        return
    ss.monday = monday
    saved = plan_file(monday)
    if saved.exists():
        load_json(saved)
    else:
        ss.days, ss.plan, ss.notes = None, {}, []
        ss.loaded_for = monday
        ss.snap = ""
        for i in range(7):
            ss[f"lock_{i}"] = False
            for slot in ("lunch", "dinner", "snack"):
                ss.pop(f"{slot}_{i}", None)
        ss.week_start = monday


def workbook_path() -> Path:
    return Path(ss.get("wb_path", str(mp.DEFAULT_WORKBOOK))).expanduser().resolve()


@st.cache_resource
def load_workbook_cached(path: str, mtime_ns: int):
    workbook = load_workbook(path)
    if Path(path).stat().st_mtime_ns != mtime_ns:
        return load_workbook(path)
    return workbook


def get_workbook(path: Path):
    return load_workbook_cached(str(path), path.stat().st_mtime_ns)


def save_wb(wb) -> bool:
    path = workbook_path()
    try:
        mp.backup_once(path)
        wb.save(path)
        return True
    except PermissionError:
        st.error("Could not save - is the workbook open in Excel? Close it and try again.")
        return False


def model_choices(current: str) -> list[str]:
    try:
        import ollama
        resp = ollama.list()
        names = [getattr(m, "model", None) or m.get("model") for m in (getattr(resp, "models", None) or resp["models"])]
        return sorted(set(names + [current]))
    except Exception:
        return [current]


@st.cache_data(ttl=60, show_spinner=False)
def cached_model_choices(current: str) -> list[str]:
    return model_choices(current)


# ------------------------------------------------------------------ boot
init_state()
cfg = ss.cfg

with st.sidebar:
    st.title("🥗 Meal Planner")
    st.text_input("Workbook", value=str(mp.DEFAULT_WORKBOOK), key="wb_path")
    if st.button("🔄 Reload workbook"):
        st.rerun()
    st.caption("Close the file in Excel before saving from here.")
    history = sorted(PLANS_DIR.glob("plan_*.json"), reverse=True) if PLANS_DIR.exists() else []
    if history:
        pick = st.selectbox("Past plans", history, format_func=lambda p: "Week of " + p.stem.replace("plan_", ""))
        if st.button("Load that plan"):
            load_json(pick)
            st.rerun()

path = workbook_path()
if not path.exists():
    st.error(f"Workbook not found: {path}")
    st.stop()
try:
    wb = get_workbook(path)
    ctx = mp.load_context(wb, cfg)
except (ValueError, KeyError) as exc:
    st.error(str(exc))
    st.stop()

if ss.loaded_for != ss.monday:
    saved_plan = plan_file(ss.monday)
    if saved_plan.exists():
        load_json(saved_plan)
        st.rerun()
    ss.loaded_for = ss.monday

status_of = {d["name"]: d["status"] for d in ctx.dishes}
dish_names = sorted({d["name"] for d in ctx.lunch_pool + ctx.dinner_pool}, key=str.lower)


def label(name: str) -> str:
    s = status_of.get(name)
    return name + (" 🆕" if s and s != "known" else "")


tab_week, tab_shop, tab_dishes, tab_pantry, tab_settings = st.tabs(
    ["📅 This week", "🛒 Shopping", "🥘 Dishes", "🧺 Pantry", "⚙️ Settings"])

st.markdown(
    """
    <style>
      [data-testid="stMetric"] {
          
          border: 1px solid #e2e8f0;
          border-radius: 12px;
          padding: 12px;
      }
      [data-testid="stDataFrame"] {
          border: 1px solid #e2e8f0;
          border-radius: 12px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------ THIS WEEK
with tab_week:
    st.date_input("Week starting (Monday)", key="week_start", on_change=change_week)
    if ss.days is None:
        st.info("No plan yet. Mark a few dishes as Known in Dishes, then generate a week.")

    locked_idx = [i for i in range(7) if ss.get(f"lock_{i}")]
    if ss.days is None:
        button_label = "Generate week"
    elif locked_idx:
        button_label = f"Regenerate (keep {len(locked_idx)} locked day(s))"
    else:
        button_label = "Regenerate week"
    if st.button(button_label, type="primary"):
        locked = ({i: {k: ss.days[i][k] for k in ("lunch", "dinner", "snack")} for i in locked_idx}
                  if ss.days else None)
        ss.log = []
        try:
            with st.spinner(f"Asking {cfg['model']} ... (a 7B model can take a minute)"):
                plan, days, notes = mp.generate(ctx, cfg, locked, log=ss.log.append)
            set_plan(days, plan, notes, ss.monday)
            save_json()
            st.rerun()
        except SystemExit as exc:
            st.error(str(exc))
    if ss.log:
        with st.expander("What the model did (retries, rule fixes)"):
            st.code("\n".join(ss.log))

    if ss.days is not None:
        for i, d in enumerate(ss.days):
            with st.container(border=True):
                head, body = st.columns([1, 4])
                head.markdown(
                    f"### {d['day']}\n"
                    f"{(ss.monday + timedelta(days=i)).strftime('%d %b')}"
                )
                head.checkbox("Lock", key=f"lock_{i}")

                c1, c2, c3 = body.columns(3)
                d["lunch"] = c1.selectbox(
                    "Lunch", dish_names, key=f"lunch_{i}",
                    format_func=label, label_visibility="collapsed",
                )
                d["dinner"] = c2.selectbox(
                    "Dinner", dish_names, key=f"dinner_{i}",
                    format_func=label, label_visibility="collapsed",
                )
                d["snack"] = c3.selectbox(
                    "Snack", ctx.snacks, key=f"snack_{i}",
                    label_visibility="collapsed",
                )

                if mp.wants_rice(d["lunch"], cfg) or mp.wants_rice(d["dinner"], cfg):
                    st.caption("🍚 Rice will be added as a side.")

                if d["reason"]:
                    st.caption(d["reason"])

        rows, warnings = mp.analyse(ss.days, ctx, cfg)
        _, greens_days, _ = mp.build_shopping(ss.days, ctx.dishes, ctx.matchers, ctx.pantry, cfg)
        counts = pd.Series([x for d in ss.days for x in (d["lunch"], d["dinner"])]).value_counts()
        n_learn = sum(1 for d in ss.days if status_of.get(d["dinner"]) != "known")
        m = st.columns(4)
        m[0].metric("Learning dinners", f"{n_learn} / 7")
        m[1].metric("Different dishes", len(counts))
        m[2].metric("Items to buy", sum(1 for r_ in rows if r_[2] == "Buy"))
        m[3].metric("Greens days", f"{greens_days} / 7")
        for d in ss.days:
            if d["lunch"] == d["dinner"]:
                st.warning(f"{d['day']}: lunch and dinner are the same dish.")
        for n in ss.notes + warnings:
            st.warning(n)
        if ss.plan.get("prep_notes"):
            with st.expander("🧑‍🍳 Prep notes"):
                for n in ss.plan["prep_notes"]:
                    st.markdown(f"- {n}")

        ideas = [s for s in (ss.plan.get("suggested_new_dishes") or []) if isinstance(s, dict) and s.get("name")]
        if ideas:
            with st.expander(f"💡 {len(ideas)} new dish idea(s) from the model"):
                for k, s in enumerate(ideas):
                    c1, c2 = st.columns([5, 1])
                    mi = s.get("main_ingredients")
                    c1.markdown(f"**{s['name']}** ({s.get('meal', 'Dinner')}) - {s.get('why', '')}  \n"
                                f"_{', '.join(mi) if isinstance(mi, list) else mi or ''}_")
                    if c2.button("Add to Dishes", key=f"idea_{k}"):
                        w2 = load_workbook(path)
                        mp.append_dish(w2, {"name": s["name"], "meal": s.get("meal", "Dinner"), "status": "Not tried",
                                            "ingredients": ", ".join(mi) if isinstance(mi, list) else str(mi or "")})
                        if save_wb(w2):
                            st.success("Added to 'What can i make'.")
                            st.rerun()

        if st.button("💾 Save plan to Excel (Meal Plan / Shopping List / Prep Notes sheets)"):
            w2 = load_workbook(path)
            mp.ensure_pantry_sheet(w2)
            mp.write_outputs(w2, ss.days, rows, ss.plan, ctx.dishes, cfg, ss.monday, ss.notes + warnings)
            if save_wb(w2):
                save_json()
                st.success(f"Saved to {path.name} (backup kept).")
        markdown = [f"# Meal plan: week of {ss.monday.strftime('%d %b %Y')}", ""]
        for i, day in enumerate(ss.days):
            dinner = day["dinner"] + (" + rice" if mp.wants_rice(day["dinner"], cfg) else "")
            if status_of.get(day["dinner"]) != "known":
                dinner += " (learning)"
            lunch = day["lunch"] + (" + rice" if mp.wants_rice(day["lunch"], cfg) else "")
            markdown.extend([
                f"## {day['day']} - {(ss.monday + timedelta(days=i)).strftime('%d %b')}",
                f"- Breakfast: {cfg['fixed_breakfast']}", f"- Lunch: {lunch}",
                f"- Dinner: {dinner}", f"- Snack: {day['snack']}",
            ])
            if day["reason"]:
                markdown.append(f"- Why: {day['reason']}")
            markdown.append("")
        st.download_button("Download / share plan", "\n".join(markdown),
                           file_name=f"meal_plan_{ss.monday.isoformat()}.md", mime="text/markdown")
        if json.dumps(ss.days) != ss.snap:
            save_json()
            ss.snap = json.dumps(ss.days)

# ------------------------------------------------------------------ SHOPPING
with tab_shop:
    if ss.days is None:
        st.info("Generate a week first - the shopping list follows the plan, including any swaps you make.")
    else:
        rows, _ = mp.analyse(ss.days, ctx, cfg)
        df = ui.shopping_df(rows)
        show = st.radio("Show", ["To buy / check", "Everything"], horizontal=True)
        if show == "To buy / check":
            df = df[df["Status"] != "In stock"]
        df = df.assign(**{"Got it": False, "Qty": None, "Unit": ""})
        st.caption(f"Cooking for {cfg['cook_for']}, buying for {cfg['buy_for']}: enter the quantity you actually buy. "
                   "Ticked rows are added to your Pantry.")
        edited = st.data_editor(
            df, hide_index=True, width="stretch", key="shop_editor",
            disabled=ui.SHOP_COLS, column_config={
                "Got it": st.column_config.CheckboxColumn(width="small"),
                "Qty": st.column_config.NumberColumn(min_value=0, width="small"),
                "Unit": st.column_config.SelectboxColumn(options=["", "kg", "g", "pcs", "l", "ml", "bunch", "pack"], width="small"),
                "Meals": st.column_config.NumberColumn(width="small")})
        c1, c2 = st.columns(2)
        if c1.button("📥 Add ticked items to Pantry"):
            bought = [{"item": r["Item"], "qty": None if pd.isna(r["Qty"]) else float(r["Qty"]), "unit": r["Unit"] or ""}
                      for _, r in edited[edited["Got it"]].iterrows()]
            if not bought:
                st.info("Tick at least one item first.")
            else:
                w2 = load_workbook(path)
                new_rows, msgs = ui.merge_bought(ui.df_to_pantry_rows(ui.pantry_df(w2)), bought, cfg["aliases"])
                mp.write_pantry(w2, new_rows)
                if save_wb(w2):
                    st.success(f"Pantry updated with {len(bought)} item(s).")
                    for m_ in msgs:
                        st.warning(m_)
                    st.rerun()
        text = ui.shopping_text(rows, f"Shopping list - week of {ss.monday.strftime('%d %b')}")
        with c2:
            st.download_button("⬇️ Download as text (to share)", text, file_name="shopping_list.txt")
        with st.expander("Copyable list"):
            st.code(text, language=None)

# ------------------------------------------------------------------ DISHES
with tab_dishes:
    orig = ui.dishes_df(ctx.dishes)
    k = st.columns(3)
    k[0].metric("Known dishes", int((orig["Status"] == "Known").sum()))
    k[1].metric("Lunch-eligible", len(ctx.lunch_pool))
    k[2].metric("Dinner-eligible", len(ctx.dinner_pool))
    st.caption("Lunch uses **Known** dishes whose meal type includes Lunch (or is blank). Every dish you mark Known "
               "gives next week's lunches more variety. Filling in **Ingredients** makes the shopping list accurate.")
    edited = st.data_editor(
        orig, width="stretch", key="dish_editor", disabled=["Dish"], height=420,
        column_config={"Ease": st.column_config.SelectboxColumn(options=ui.EASE),
                       "Meal": st.column_config.SelectboxColumn(options=ui.MEALS),
                       "Status": st.column_config.SelectboxColumn(options=ui.STATUS),
                       "Ingredients": st.column_config.TextColumn(width="large")})
    changes = ui.diff_dishes(orig, edited)
    if st.button(f"💾 Save dish changes ({len(changes)} cell(s))", disabled=not changes):
        w2 = load_workbook(path)
        mp.update_dish_cells(w2, changes)
        if save_wb(w2):
            st.success("Saved.")
            st.rerun()
    with st.expander("➕ Add a new dish"):
        with st.form("add_dish", clear_on_submit=True):
            n = st.text_input("Dish name")
            ing = st.text_input("Ingredients (comma separated)")
            f1, f2, f3 = st.columns(3)
            ease, meal, stat = f1.selectbox("Ease", ui.EASE[1:]), f2.selectbox("Meal", ui.MEALS[1:], index=3), f3.selectbox("Status", ui.STATUS, index=3)
            link = st.text_input("Recipe link (optional)")
            if st.form_submit_button("Add dish") and n.strip():
                w2 = load_workbook(path)
                mp.append_dish(w2, {"name": n.strip(), "ingredients": ing, "ease": ease, "meal": meal,
                                    "status": None if stat == "Unknown" else stat, "link": link})
                if save_wb(w2):
                    st.rerun()
    if ctx.excluded:
        with st.expander(f"Why {len(ctx.excluded)} dish(es) can't be planned"):
            for name, why in ctx.excluded:
                st.markdown(f"- **{name}** - {why}")

# ------------------------------------------------------------------ PANTRY
with tab_pantry:
    st.caption("What's at home. Anything with a quantity above 0 is treated as in stock and won't be marked 'Buy'.")
    pdf = ui.pantry_df(wb)
    pedit = st.data_editor(pdf, num_rows="dynamic", width="stretch", hide_index=True, key="pantry_editor",
                           column_config={"Available quantity": st.column_config.NumberColumn(min_value=0)})
    if st.button("💾 Save pantry"):
        w2 = load_workbook(path)
        mp.write_pantry(w2, ui.df_to_pantry_rows(pedit))
        if save_wb(w2):
            st.success("Pantry saved.")
            st.rerun()

# ------------------------------------------------------------------ SETTINGS
with tab_settings:
    with st.form("settings"):
        c1, c2 = st.columns(2)
        models = cached_model_choices(cfg["model"])
        model = c1.selectbox("Ollama model", models, index=models.index(cfg["model"]))
        temp = c2.slider("Creativity (temperature)", 0.0, 1.0, float(cfg["temperature"]), 0.05)
        breakfast = st.text_input("Fixed breakfast", cfg["fixed_breakfast"])
        c3, c4, c5 = st.columns(3)
        cook_for = c3.number_input("Cook for (people)", 1, 10, int(cfg["cook_for"]))
        buy_for = c4.number_input("Buy for (people)", 1, 10, int(cfg["buy_for"]))
        min_learn = c5.slider("Min learning dinners / week", 0, 7, int(cfg["min_learning_dinners"]))
        c6, c7, c8 = st.columns(3)
        rep_l = c6.number_input("Max repeats of a lunch dish", 1, 7, int(cfg["max_repeats_lunch"]))
        rep_d = c7.number_input("Max repeats of a dinner dish", 1, 7, int(cfg["max_repeats_dinner"]))
        rice = c8.checkbox("Add rice as a side to dals/curries", bool(cfg["serve_rice"]))
        can = st.text_input("Can cook (comma separated)", ", ".join(cfg["can_cook"]))
        cannot = st.text_input("Can't cook yet (excluded from plans)", ", ".join(cfg["cannot_cook_yet"]))
        staples = st.text_input("Weekly staples for breakfast", ", ".join(cfg["weekly_staples"]))
        ignore = st.text_area("Never put these on the shopping list (words, comma separated)", ", ".join(cfg["ignore_ingredients"]))
        diet = st.text_area("Diet", cfg["diet"])
        health = st.text_area("Health context given to the model", cfg["health_context"])
        if st.form_submit_button("Save settings"):
            split = lambda t: [x.strip() for x in t.split(",") if x.strip()]
            cfg.update(model=model, temperature=temp, fixed_breakfast=breakfast, cook_for=cook_for, buy_for=buy_for,
                       min_learning_dinners=min_learn, max_repeats_lunch=rep_l, max_repeats_dinner=rep_d, serve_rice=rice,
                       can_cook=split(can), cannot_cook_yet=split(cannot), weekly_staples=split(staples),
                       ignore_ingredients=split(ignore), diet=diet, health_context=health)
            mp.CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
            st.success("Saved to config.json.")
            st.rerun()