"""Personalised weekly meal planner. Run with: streamlit run app.py"""
from __future__ import annotations
import json
from datetime import date, timedelta
from io import BytesIO
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import meal_planner as mp
import ui_logic as ui
from services import supabase_service as db
from services import research_service as research

st.set_page_config(page_title="My Meal Planner", page_icon="🥗", layout="wide")
ss = st.session_state


def auth_screen():
    st.title("🥗 My Meal Planner")
    st.caption("Your meals, pantry and preferences — saved privately to your account.")
    login_tab, signup_tab = st.tabs(["Log in", "Create account"])
    with login_tab:
        with st.form("login_form"):
            email = st.text_input("Email", key="login_email")
            password = st.text_input("Password", type="password", key="login_password")
            submitted = st.form_submit_button("Log in", type="primary", use_container_width=True)
        if submitted:
            try:
                client = db.get_client()
                response = db.sign_in(client, email.strip(), password)
                if not response.user:
                    st.error("Login did not return a user. Check your credentials and email confirmation.")
                else:
                    ss.supabase_client = client
                    ss.user_id = response.user.id
                    ss.user_email = response.user.email
                    st.rerun()
            except Exception as exc:
                st.error(f"Could not log in: {exc}")
    with signup_tab:
        with st.form("signup_form"):
            email = st.text_input("Email address", key="signup_email")
            password = st.text_input("Password (at least 8 characters)", type="password", key="signup_password")
            confirm = st.text_input("Confirm password", type="password")
            submitted = st.form_submit_button("Create account", type="primary", use_container_width=True)
        if submitted:
            if len(password) < 8:
                st.error("Use a password with at least 8 characters.")
            elif password != confirm:
                st.error("The passwords do not match.")
            else:
                try:
                    client = db.get_client()
                    response = db.sign_up(client, email.strip(), password)
                    if response.session and response.user:
                        ss.supabase_client = client
                        ss.user_id = response.user.id
                        ss.user_email = response.user.email
                        st.rerun()
                    else:
                        st.success("Account created. Check your email to confirm the account, then log in.")
                except Exception as exc:
                    st.error(f"Could not create account: {exc}")


def workbook_setup():
    st.title("📘 Set up your workbook")
    st.write("Upload your Excel workbook, or restore the workbook you last saved to your account.")
    client, uid = ss.supabase_client, ss.user_id
    try:
        saved = db.fetch_workbook(client, uid)
    except Exception as exc:
        saved = None
        st.error(f"Could not check Supabase Storage: {exc}")
    if saved:
        meta, content = saved
        st.info(f"Saved workbook: **{meta['file_name']}** · Last saved: {meta.get('updated_at', 'unknown')}")
        if st.button("Use my saved workbook", type="primary"):
            ss.workbook_bytes = content
            ss.workbook_name = meta["file_name"]
            ss.workbook_ready = True
            st.rerun()
    else:
        st.info("No workbook is saved to your account yet. Upload one to get started.")
    uploaded = st.file_uploader("Upload workbook (.xlsx)", type=["xlsx"])
    if uploaded and st.button("Upload and continue", type="primary"):
        try:
            content = uploaded.getvalue()
            # Validate that it is a readable Excel workbook before uploading.
            test_wb = load_workbook(BytesIO(content))
            required = {mp.sheet_key("Grocery"), mp.sheet_key("What can i make"), mp.sheet_key("snacks")}
            found = {mp.sheet_key(x) for x in test_wb.sheetnames}
            missing = required - found
            if missing:
                st.error("Workbook is missing required sheets: " + ", ".join(sorted(missing)))
            else:
                db.upload_workbook(client, uid, uploaded.name, content)
                ss.workbook_bytes = content
                ss.workbook_name = uploaded.name
                ss.workbook_ready = True
                st.success("Workbook uploaded securely.")
                st.rerun()
        except Exception as exc:
            st.error(f"Workbook upload failed: {exc}")


def load_user_workbook():
    wb = load_workbook(BytesIO(ss.workbook_bytes))
    cfg = mp.load_config(wb)
    # The database pantry is the account-level source of truth; mirror it into the workbook.
    try:
        pantry = db.fetch_pantry(ss.supabase_client, ss.user_id)
        if pantry:
            mp.write_pantry(wb, pantry)
    except Exception as exc:
        st.warning(f"Could not load your saved pantry: {exc}")
    ctx = mp.load_context(wb, cfg)
    return wb, cfg, ctx


def save_workbook(wb):
    stream = BytesIO()
    wb.save(stream)
    content = stream.getvalue()
    db.upload_workbook(ss.supabase_client, ss.user_id, ss.workbook_name, content)
    ss.workbook_bytes = content


def set_plan(days, plan, notes, monday):
    ss.days, ss.plan, ss.notes, ss.monday = days, plan, notes, monday
    ss.snap = json.dumps(days)
    for i, day in enumerate(days):
        for slot in ("lunch", "dinner", "snack"):
            ss[f"{slot}_{i}"] = day[slot]


def load_saved_plan(monday):
    record = db.fetch_plan(ss.supabase_client, ss.user_id, monday.isoformat())
    if record:
        set_plan(record["days"], record.get("plan_data") or {}, record.get("notes") or [], monday)
        return True
    ss.days, ss.plan, ss.notes = None, {}, []
    ss.monday = monday
    ss.snap = ""
    return False


if "user_id" not in ss or "supabase_client" not in ss:
    auth_screen()
    st.stop()

with st.sidebar:
    st.title("🥗 My Meal Planner")
    st.caption(ss.get("user_email", "Signed in"))
    if st.button("Change workbook", use_container_width=True):
        ss.workbook_ready = False
        ss.pop("workbook_bytes", None)
        st.rerun()
    if st.button("Log out", use_container_width=True):
        try:
            db.sign_out(ss.supabase_client)
        finally:
            for key in ("user_id", "user_email", "supabase_client", "workbook_ready", "workbook_bytes", "workbook_name", "days", "plan", "notes", "rules"):
                ss.pop(key, None)
            st.rerun()

if not ss.get("workbook_ready"):
    workbook_setup()
    st.stop()

try:
    wb, cfg, ctx = load_user_workbook()
except Exception as exc:
    st.error(f"Could not read your workbook: {exc}")
    if st.button("Choose another workbook"):
        ss.workbook_ready = False
        st.rerun()
    st.stop()

if "rules" not in ss:
    try:
        ss.rules = db.fetch_rules(ss.supabase_client, ss.user_id)
    except Exception as exc:
        ss.rules = []
        st.warning(f"Could not load saved food rules: {exc}")
if "monday" not in ss:
    ss.monday = mp.plan_monday()
    ss.week_start = ss.monday
    ss.days, ss.plan, ss.notes, ss.snap = None, {}, [], ""

status_of = {d["name"]: d["status"] for d in ctx.dishes}
lunch_names = sorted({d["name"] for d in ctx.lunch_pool}, key=str.lower)
dinner_names = sorted({d["name"] for d in ctx.dinner_pool}, key=str.lower)

tab_week, tab_shop, tab_prep, tab_nutrition, tab_rules, tab_prices = st.tabs([
    "📅 Weekly Planner", "🛒 Shopping + Pantry", "🍱 Meal Prep", "🥦 Nutrition Discovery", "⚙️ My Preferences", "💸 Price Research"
])

# WEEKLY PLANNER
with tab_week:
    st.subheader("Your weekly meals")
    chosen_date = st.date_input("Week starting", value=ss.get("week_start", ss.monday))
    monday = chosen_date - timedelta(days=chosen_date.weekday())
    if monday != ss.monday:
        load_saved_plan(monday)
        ss.week_start = monday
    active_rules = [r for r in ss.rules if r.get("active", True)]
    fixed_rules = [r for r in active_rules if r.get("rule_type") == "fixed_meal"]
    if fixed_rules:
        st.caption("Permanent fixed meals: " + "; ".join(f"{r['day']} {r['meal']}: {r.get('dish') or r.get('value')}" for r in fixed_rules))
    locked_idx = [i for i in range(7) if ss.get(f"lock_{i}")]
    label = "Generate week" if ss.days is None else (f"Regenerate (keep {len(locked_idx)} day(s)" if locked_idx else "Regenerate week")
    if st.button(label, type="primary"):
        locked = ({i: {k: ss.days[i][k] for k in ("lunch", "dinner", "snack")} for i in locked_idx} if ss.days else None)
        try:
            with st.spinner("Applying your workbook, nutrition priorities and permanent rules..."):
                plan, days, notes = mp.generate(ctx, cfg, locked, active_rules)
            set_plan(days, plan, notes, monday)
            db.save_plan(ss.supabase_client, ss.user_id, monday.isoformat(), days, plan, notes)
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))
    if ss.days is None:
        st.info("Generate a week to see your meals. Your workbook is the source of recipes and grocery items.")
    else:
        st.caption(f"Breakfast every day: {cfg['fixed_breakfast']}")
        st.markdown("**Lunch · Dinner · Snack**")
        for i, day in enumerate(ss.days):
            with st.container(border=True):
                cday, clock = st.columns([5, 1])
                cday.markdown(f"**{day['day']} · {(ss.monday + timedelta(days=i)).strftime('%d %b')}**")
                clock.checkbox("Keep", key=f"lock_{i}", help="Keep this day's meals when regenerating")
                a, b, c = st.columns(3)
                adjacent = set()
                for ni in (i - 1, i + 1):
                    if 0 <= ni < len(ss.days):
                        adjacent.update((ss.days[ni]["lunch"], ss.days[ni]["dinner"]))
                fixed_lunch = next((r.get("dish") or r.get("value") for r in active_rules
                                    if r.get("rule_type") == "fixed_meal" and str(r.get("day", "")).lower() == day["day"].lower()
                                    and str(r.get("meal", "")).lower() == "lunch"), None)
                fixed_dinner = next((r.get("dish") or r.get("value") for r in active_rules
                                     if r.get("rule_type") == "fixed_meal" and str(r.get("day", "")).lower() == day["day"].lower()
                                     and str(r.get("meal", "")).lower() == "dinner"), None)
                if fixed_lunch:
                    day["lunch"] = fixed_lunch
                    ss[f"lunch_{i}"] = fixed_lunch
                    a.selectbox("Lunch (fixed rule)", [fixed_lunch], key=f"lunch_{i}", disabled=True)
                else:
                    day["lunch"] = a.selectbox("Lunch", ui.meal_options(lunch_names, day["lunch"], day["dinner"], adjacent), key=f"lunch_{i}")
                if fixed_dinner:
                    day["dinner"] = fixed_dinner
                    ss[f"dinner_{i}"] = fixed_dinner
                    b.selectbox("Dinner (fixed rule)", [fixed_dinner], key=f"dinner_{i}", disabled=True)
                else:
                    day["dinner"] = b.selectbox("Dinner", ui.meal_options(dinner_names, day["dinner"], day["lunch"], adjacent), key=f"dinner_{i}")
                day["snack"] = c.selectbox("Snack", ctx.snacks or ["No snack listed"], key=f"snack_{i}")
                if day.get("reason"):
                    st.caption(day["reason"])
        rows, warnings = mp.analyse(ss.days, ctx, cfg)
        m1, m2 = st.columns(2)
        m1.metric("Different main dishes", len({x for d in ss.days for x in (d["lunch"], d["dinner"])}))
        m2.metric("Items to buy", sum(1 for r in rows if r[2] == "Buy"))
        for warning in ss.notes + warnings:
            st.warning(warning)
        save_col, excel_col = st.columns(2)
        if save_col.button("Save current plan", type="primary"):
            db.save_plan(ss.supabase_client, ss.user_id, ss.monday.isoformat(), ss.days, ss.plan, ss.notes + warnings)
            st.success("Plan saved to your Supabase account.")
        if excel_col.button("Write plan + shopping sheets into workbook"):
            rows_for_excel, warnings_for_excel = mp.analyse(ss.days, ctx, cfg)
            mp.write_outputs(wb, ss.days, rows_for_excel, ss.plan, ctx.dishes, cfg, ss.monday, ss.notes + warnings_for_excel)
            save_workbook(wb)
            st.success("Meal Plan, Shopping List and Meal Prep sheets saved to your Supabase workbook.")
        markdown = [f"# Meal plan: week of {ss.monday:%d %b %Y}", ""]
        for i, day in enumerate(ss.days):
            markdown += [f"## {day['day']} - {(ss.monday + timedelta(days=i)):%d %b}", f"- Breakfast: {cfg['fixed_breakfast']}", f"- Lunch: {day['lunch']}", f"- Dinner: {day['dinner']}", f"- Snack: {day['snack']}", ""]
        st.download_button("Download weekly plan", "\n".join(markdown), file_name=f"meal_plan_{ss.monday}.md")

# SHOPPING + PANTRY
with tab_shop:
    st.subheader("Shopping list")
    if ss.days is None:
        st.info("Generate a weekly plan first.")
    else:
        rows, _ = mp.analyse(ss.days, ctx, cfg)
        df = ui.shopping_df(rows)
        view = st.radio("Show", ["To buy / check", "Everything"], horizontal=True)
        if view == "To buy / check":
            df = df[df.Status != "In stock"]
        df = df.assign(**{"Got it": False, "Qty": None, "Unit": ""})
        edited = st.data_editor(df, hide_index=True, use_container_width=True, disabled=ui.SHOP_COLS,
            column_config={"Got it": st.column_config.CheckboxColumn(), "Qty": st.column_config.NumberColumn(min_value=0),
                "Unit": st.column_config.SelectboxColumn(options=["", "kg", "g", "pcs", "l", "ml", "bunch", "pack"])})
        if st.button("Add ticked items to pantry"):
            bought = [{"item": r["Item"], "qty": None if pd.isna(r["Qty"]) else float(r["Qty"]), "unit": r["Unit"] or ""}
                      for _, r in edited[edited["Got it"]].iterrows()]
            new_rows, messages = ui.merge_bought(ui.df_to_pantry_rows(ui.pantry_df(wb)), bought, cfg["aliases"])
            # Use the workbook representation so item aliases and units remain consistent.
            mp.write_pantry(wb, new_rows)
            save_workbook(wb)
            db.save_pantry(ss.supabase_client, ss.user_id, new_rows)
            for message in messages:
                st.warning(message)
            st.success(f"Updated pantry with {len(bought)} purchase(s).")
            st.rerun()
        st.download_button("Download shopping list", ui.shopping_text(rows, "Shopping list"), file_name="shopping_list.txt")
    st.divider()
    st.subheader("Pantry")
    pantry_df = ui.pantry_df(wb)
    pantry_edit = st.data_editor(pantry_df, num_rows="dynamic", hide_index=True, use_container_width=True,
        column_config={"Available quantity": st.column_config.NumberColumn(min_value=0)})
    if st.button("Save pantry"):
        pantry_rows = ui.df_to_pantry_rows(pantry_edit)
        mp.write_pantry(wb, pantry_rows)
        save_workbook(wb)
        db.save_pantry(ss.supabase_client, ss.user_id, pantry_rows)
        st.success("Pantry saved to your workbook and Supabase.")

# MEAL PREP
with tab_prep:
    st.subheader("Meal preparation")
    if ss.days is None:
        st.info("Generate a weekly plan first.")
    else:
        prep_rows, _ = mp.analyse(ss.days, ctx, cfg)
        prep = mp.meal_prep_text(ss.days, ctx.dishes, prep_rows, cfg, ss.monday)
        st.markdown(prep)
        st.download_button("Download meal prep report", prep, file_name=f"meal_prep_{ss.monday}.md")

# NUTRITION DISCOVERY
with tab_nutrition:
    st.subheader("Nutrition discovery")
    st.warning("This is food research, not diagnosis or treatment. Enter lab units and reference ranges exactly as shown on your report; discuss abnormal results and targets with a qualified clinician.")
    with st.form("nutrition_form"):
        nutrient = st.text_input("Nutrient / deficiency", placeholder="e.g. iron, vitamin B12, vitamin D")
        c1, c2 = st.columns(2)
        reported = c1.text_input("Your reported value (optional)")
        unit = c2.text_input("Unit (as printed on report)", placeholder="e.g. ng/mL")
        reference = st.text_input("Lab reference range (optional)")
        target = st.text_input("Clinician-recommended target (optional)")
        notes = st.text_area("Notes (optional)")
        submitted = st.form_submit_button("Save nutrient profile")
    if submitted:
        if not nutrient.strip():
            st.error("Enter a nutrient name.")
        else:
            existing = db.fetch_nutrition_profile(ss.supabase_client, ss.user_id)
            existing = [x for x in existing if x["nutrient"].lower() != nutrient.strip().lower()]
            existing.append({"nutrient": nutrient.strip(), "reported_value": reported or None, "unit": unit or None,
                             "reference_range": reference or None, "target": target or None, "notes": notes or None})
            db.save_nutrition_profile(ss.supabase_client, ss.user_id, existing)
            st.success("Saved to your private nutrition profile.")
    try:
        profiles = db.fetch_nutrition_profile(ss.supabase_client, ss.user_id)
    except Exception:
        profiles = []
    if profiles:
        st.markdown("**Saved profile**")
        st.dataframe(pd.DataFrame(profiles), hide_index=True, use_container_width=True)
    st.divider()
    food_query = st.text_input("Research foods and Indian meal ideas for", value=ss.get("nutrition_search", ""), key="nutrition_search")
    if st.button("Research food sources and recipes"):
        if not food_query.strip():
            st.warning("Enter a nutrient or food.")
        else:
            try:
                with st.spinner("Searching nutrition references and recipe ideas..."):
                    ss.nutrition_results = research.search_nutrition(food_query.strip())
                    ss.recipe_results = research.search_recipes(food_query.strip())
            except Exception as exc:
                st.error(str(exc))
    if food_query.strip():
        matching_priorities = [p for p in ctx.nutrition_priorities if food_query.strip().lower() in p["name"].lower()]
        if matching_priorities:
            terms = [term for priority in matching_priorities for term in priority["foods"]]
            matching_dishes = [d["name"] for d in ctx.dishes
                               if any(term in (d["name"] + " " + d["ingredients"]).lower() for term in terms)]
            st.markdown("**Meals already in your workbook that match its nutrient food keywords**")
            st.write(matching_dishes or "No existing dish matched the workbook's nutrient keywords. You can add a researched dish below.")
    for title, key in (("Nutrition research", "nutrition_results"), ("Recipe ideas", "recipe_results")):
        if ss.get(key):
            st.markdown(f"**{title}**")
            for result in ss[key]:
                st.markdown(f"- [{result['title']}]({result['url']}) — {result['description']}")
    st.caption("To add a researched dish, enter its details below. The dish is appended to your uploaded workbook and saved to Supabase.")
    with st.form("add_dish_form"):
        dish = st.text_input("Dish name")
        ingredients = st.text_area("Ingredients (comma-separated)")
        meal_type = st.selectbox("Meal type", ["Lunch", "Dinner", "Lunch / Dinner"])
        recipe_link = st.text_input("Recipe link (optional)")
        if st.form_submit_button("Add dish to my workbook"):
            if not dish.strip():
                st.error("Enter a dish name.")
            else:
                ws = mp.get_sheet(wb, "What can i make")
                headers = [c.value for c in ws[1]]
                values = [None] * len(headers)
                mapping = {"Dish": dish.strip(), "Ingredients": ingredients, "B / L / D": meal_type, "Link": recipe_link, "Cooking status": "Not tried", "Ease (E/M/D)": "Medium"}
                for h, value in mapping.items():
                    col = mp.find_col(headers, [h])
                    if col is not None:
                        values[col] = value
                ws.append(values)
                save_workbook(wb)
                st.success(f"Added {dish} to your workbook. Reload the app to refresh dish options.")

# PREFERENCES / RULES
with tab_rules:
    st.subheader("Permanent food rules")
    st.write("Hard rules are enforced during generation. Soft preferences influence meal selection but do not guarantee a particular dish.")
    with st.form("add_rule_form"):
        rule_type = st.selectbox("Rule type", ["fixed_meal", "avoid_dish", "avoid_ingredient", "prefer_ingredient"],
            format_func=lambda x: {"fixed_meal": "Always eat a specific dish", "avoid_dish": "Never select a dish", "avoid_ingredient": "Avoid an ingredient", "prefer_ingredient": "Prefer an ingredient"}[x])
        day = meal = dish_value = rule_value = None
        if rule_type == "fixed_meal":
            c1, c2 = st.columns(2)
            day = c1.selectbox("Day", mp.DAYS)
            meal = c2.selectbox("Meal", ["Lunch", "Dinner"])
            eligible_names = lunch_names if meal == "Lunch" else dinner_names
            dish_value = st.selectbox("Dish", eligible_names)
        elif rule_type == "avoid_dish":
            dish_value = st.selectbox("Dish", sorted({d["name"] for d in ctx.dishes}, key=str.lower))
        else:
            rule_value = st.text_input("Ingredient")
        if st.form_submit_button("Add permanent rule"):
            rule = {"rule_type": rule_type, "day": day, "meal": meal.lower() if meal else None,
                    "dish": dish_value, "value": rule_value, "active": True}
            ss.rules.append(rule)
            try:
                db.save_rules(ss.supabase_client, ss.user_id, ss.rules)
                st.success("Rule saved. It will apply to future generations.")
            except Exception as exc:
                ss.rules.pop()
                st.error(f"Could not save rule: {exc}")
    if ss.rules:
        rules_df = pd.DataFrame(ss.rules)
        edited_rules = st.data_editor(rules_df, hide_index=True, use_container_width=True,
            column_config={"active": st.column_config.CheckboxColumn("Active")})
        c1, c2 = st.columns(2)
        if c1.button("Save rule changes"):
            ss.rules = edited_rules.to_dict("records")
            db.save_rules(ss.supabase_client, ss.user_id, ss.rules)
            st.success("Preferences updated.")
        if c2.button("Delete all rules"):
            ss.rules = []
            db.save_rules(ss.supabase_client, ss.user_id, [])
            st.rerun()
    else:
        st.info("No permanent rules yet. Try: Tuesday lunch → Curd Rice.")

# PRICE RESEARCH
with tab_prices:
    st.subheader("Grocery price research")
    st.caption("Search-index results are leads, not guaranteed current prices or local availability. Open the retailer listing to verify price, pack size, seller, fees and delivery area before buying.")
    item = st.text_input("Grocery item", placeholder="e.g. tomatoes 500 g")
    area = st.text_input("Delivery area / city", value="Hyderabad")
    if st.button("Compare retailer search results", type="primary"):
        if not item.strip():
            st.warning("Enter an item to research.")
        else:
            try:
                ss.price_results = research.search_grocery_item(item.strip(), area.strip() or "India")
                ss.price_item = item.strip()
            except Exception as exc:
                st.error(str(exc))
                ss.price_results = None
    if ss.get("price_results"):
        for platform, results in ss.price_results.items():
            with st.expander(platform, expanded=True):
                if isinstance(results, dict) and "error" in results:
                    st.caption(results["error"])
                elif results:
                    for result in results:
                        price_text = (" · price mention: ₹" + ", ₹".join(result["price_mentions"])) if result["price_mentions"] else " · no price detected in snippet"
                        st.markdown(f"[{result['title']}]({result['url']}){price_text}")
                        st.caption(result["description"])
                        st.caption("Search-index snippet — verify the listing directly; not a confirmed offer.")
                else:
                    st.caption("No indexed result found.")
    else:
        st.markdown("**Open retailer-specific searches**")
        if item.strip():
            for platform, domain in research.PLATFORMS.items():
                query = quote_plus(f"{item} {area}")
                st.markdown(f"- [{platform} search](https://www.google.com/search?q=site%3A{quote_plus(domain)}+{query})")

st.divider()
st.caption(f"Workbook: {ss.workbook_name} · Signed in as {ss.user_email}. Your workbook, rules, pantry and weekly plans are stored in your Supabase account.")
