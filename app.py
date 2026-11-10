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
    ss.setdefault("days", None)
    ss.setdefault("plan", {})
    ss.setdefault("notes", [])
    ss.setdefault("monday", mp.plan_monday())
    ss.setdefault("week_start", ss.monday)
    ss.setdefault("loaded_for", None)
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
        wb.save(path)
        return True
    except PermissionError:
        st.error("Could not save - is the workbook open in Excel? Close it and try again.")
        return False


# ------------------------------------------------------------------ boot
init_state()

with st.sidebar:
    st.title("🥗 Meal Planner")
    st.text_input("Workbook", value=str(mp.DEFAULT_WORKBOOK), key="wb_path")
    if st.button("🔄 Reload workbook"):
        st.rerun()
    st.caption("Close the workbook in Excel before saving changes here.")
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
    cfg = mp.load_config(wb)
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
lunch_names = sorted({d["name"] for d in ctx.lunch_pool}, key=str.lower)
dinner_names = sorted({d["name"] for d in ctx.dinner_pool}, key=str.lower)


def label(name: str) -> str:
    s = status_of.get(name)
    return name + (" 🆕" if s and s != "known" else "")


tab_week, tab_shop, tab_prep = st.tabs(["📅 This week", "🛒 Shopping + Pantry", "🍱 Meal prep"])

# ------------------------------------------------------------------ THIS WEEK
with tab_week:
    st.date_input("Week starting (Monday)", key="week_start", on_change=change_week)
    if ss.days is None:
        st.info("No plan yet. Add familiar dishes to the 'What can i make' sheet in your workbook, then generate a week.")

    st.caption("Meal selection favors easy, familiar dishes and the nutrient priorities in your workbook. Edit the workbook to change your dish options.")

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
        try:
            with st.spinner("Building a balanced week from your workbook dishes..."):
                plan, days, notes = mp.generate(ctx, cfg, locked)
            set_plan(days, plan, notes, ss.monday)
            save_json()
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))

    if ss.days is not None:
        st.caption(f"Breakfast every day: {cfg['fixed_breakfast']}")
        headers = st.columns(3)
        for col, title in zip(headers, ("Lunch", "Dinner", "Snack")):
            col.markdown(f"**{title}**")
        for i, d in enumerate(ss.days):
            with st.container(border=True):
                day_col, lock_col = st.columns([5, 1])
                day_col.markdown(f"**{d['day']} · {(ss.monday + timedelta(days=i)).strftime('%d %b')}**")
                lock_col.checkbox("Keep", key=f"lock_{i}", label_visibility="collapsed",
                                  help="Keep this day's meals when regenerating")
                lunch_col, dinner_col, snack_col = st.columns(3)
                adjacent = set()
                for neighbor_i in (i - 1, i + 1):
                    if 0 <= neighbor_i < len(ss.days):
                        adjacent.update((ss.days[neighbor_i]["lunch"], ss.days[neighbor_i]["dinner"]))
                d["lunch"] = lunch_col.selectbox(
                    "Lunch", ui.meal_options(lunch_names, d["lunch"], d["dinner"], adjacent), key=f"lunch_{i}",
                    format_func=label, label_visibility="collapsed",
                )
                d["dinner"] = dinner_col.selectbox(
                    "Dinner", ui.meal_options(dinner_names, d["dinner"], d["lunch"], adjacent), key=f"dinner_{i}",
                    format_func=label, label_visibility="collapsed",
                )
                d["snack"] = snack_col.selectbox(
                    "Snack", ctx.snacks, key=f"snack_{i}",
                    label_visibility="collapsed",
                )

                if mp.wants_rice(d["lunch"], cfg) or mp.wants_rice(d["dinner"], cfg):
                    st.caption("🍚 Rice will be added as a side.")

                if d["reason"]:
                    st.caption(d["reason"])

        rows, warnings = mp.analyse(ss.days, ctx, cfg)
        counts = pd.Series([x for d in ss.days for x in (d["lunch"], d["dinner"])]).value_counts()
        m = st.columns(2)
        m[0].metric("Different main dishes", len(counts))
        m[1].metric("Items to buy", sum(1 for r_ in rows if r_[2] == "Buy"))
        for d in ss.days:
            if d["lunch"] == d["dinner"]:
                st.warning(f"{d['day']}: lunch and dinner are the same dish.")
        for i in range(1, len(ss.days)):
            previous = {ss.days[i - 1]["lunch"], ss.days[i - 1]["dinner"]}
            repeated = previous & {ss.days[i]["lunch"], ss.days[i]["dinner"]}
            if repeated:
                st.warning(f"{ss.days[i]['day']} repeats a main dish from the previous day: {', '.join(sorted(repeated))}.")
        for n in ss.notes + warnings:
            st.warning(n)
        # if ss.plan.get("prep_notes"):
        #     with st.expander("🧑‍🍳 Prep notes"):
        #         for n in ss.plan["prep_notes"]:
        #             st.markdown(f"- {n}")

        if st.button("💾 Save plan to Excel (Meal Plan / Shopping List)"):
            w2 = load_workbook(path)
            mp.ensure_pantry_sheet(w2)
            mp.write_outputs(w2, ss.days, rows, ss.plan, ctx.dishes, cfg, ss.monday, ss.notes + warnings)
            if save_wb(w2):
                save_json()
                st.success(f"Saved to {path.name}.")
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
    st.subheader("Shopping list")
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

    st.divider()
    st.subheader("Pantry")
    st.caption("Items with a quantity above zero are treated as in stock.")
    pdf = ui.pantry_df(wb)
    pedit = st.data_editor(pdf, num_rows="dynamic", width="stretch", hide_index=True, key="pantry_editor",
                           column_config={"Available quantity": st.column_config.NumberColumn(min_value=0)})
    if st.button("💾 Save pantry"):
        w2 = load_workbook(path)
        mp.write_pantry(w2, ui.df_to_pantry_rows(pedit))
        if save_wb(w2):
            st.success("Pantry saved.")
            st.rerun()

# ------------------------------------------------------------------ MEAL PREP
with tab_prep:
    if ss.days is None:
        st.info("Generate a week first to see the meal-prep schedule and pantry gaps.")
    else:
        prep_rows, _ = mp.analyse(ss.days, ctx, cfg)
        prep_text = mp.meal_prep_text(ss.days, ctx.dishes, prep_rows, cfg, ss.monday)
        st.markdown(prep_text)
        st.download_button(
            "Download meal prep and missing ingredients",
            prep_text,
            file_name=f"meal_prep_{ss.monday.isoformat()}.md",
            mime="text/markdown",
        )
