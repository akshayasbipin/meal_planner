# Weekly Meal Planner

A local Streamlit app that builds a weekly plan and shopping list from an Excel workbook. Meal selection is randomized and rule-based; no model, internet connection, or Ollama installation is needed.

## Start

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL printed by Streamlit. The default workbook is `Grocery.xlsx`; change its path in the sidebar to use another workbook. Close the workbook in Excel before saving changes from the app.

## Workbook Setup

Sheet names are case-insensitive; spaces and underscores are interchangeable.

| Sheet | Contents |
| --- | --- |
| `Grocery` | Category names in row 1, with grocery items beneath each category. Add foods here to categorize shopping-list items. |
| `What can i make` | One dish per row. `Dish` (or `Recipe` / `Name`) is required. Add `Ingredients`, `Ease (E/M/D)`, `B / L / D`, and `Cooking status` to improve planning. `Link` and serving size are optional. |
| `snacks` | Snack names in the first column below the header. |
| `Pantry` | `Item`, `Available quantity`, and `Unit`. Created if missing. A positive quantity counts as in stock. |
| `Planner Settings` | `Setting` and `Value` columns. Controls repeat caps, fixed breakfast, servings, rice, cookable dishes, staples, ingredient exclusions, and aliases. |
| `required` | A nutrient table with `Nutrient`, `Foods to include`, and `Weekly target servings` columns. Its header can follow introductory notes. These workbook rows drive food matching and weekly selection. |

Edit dishes directly in Excel. Use meal values such as `Lunch`, `Dinner`, or `Lunch / Dinner`; a blank meal type is flexible. Mark familiar dishes `Known`, and set difficulty to Easy or Medium for lunch eligibility. Dishes listed in `cannot_cook_yet` in `Planner Settings` are excluded. List ingredients separated by commas so shopping and nutrient priorities can match them.

## Weekly Plan

Choose a week starting date and select **Generate week**. Breakfast is the fixed setting; lunch and dinner are drawn from their eligible workbook pools, and snacks are picked from the `snacks` sheet. Selection favors Known and easier dishes and uses the food terms and weekly targets in the workbook's `required` sheet. Edit those rows to change which foods and nutrient priorities the planner matches.

Lunch and dinner differ on the same day. A main dish is not repeated in either slot on the next day, and configurable weekly repeat caps still apply. Use the **Keep** checkbox to preserve a day while regenerating. Manual dropdown choices also block same-day and adjacent-day repeats; existing saved conflicts remain selectable so they can be corrected and are flagged. Plans save to `plans/plan_<Monday>.json` as you edit. Use **Past plans** to reopen a saved week.

The nutrition matching is a simple keyword preference, not nutrient analysis or medical advice. Food does not replace clinician-directed care for folate, vitamin D, or B12 deficiencies.

## Shopping + Pantry

The shopping list updates with the current plan. **To buy / check** hides items already in the pantry; **Everything** shows the full list. Enter the quantity and unit purchased, tick **Got it**, then select **Add ticked items to Pantry**. Compatible quantities are added; different units replace the old quantity with a warning. The pantry table is directly below the shopping list and can also be edited and saved manually. Items absent from the Grocery sheet are marked for checking.

Download the list as text or expand **Copyable list**. The planner calculates shopping needs from workbook ingredients, pantry contents, and configured weekly breakfast staples.

## Meal Prep

The **Meal prep** tab gives a day-by-day meal and ingredient schedule, compares required shopping items with the Pantry sheet, and separates confirmed items to buy from items that need a manual check because they are missing from the Grocery sheet. Dishes without workbook ingredients are flagged. Download the same report as Markdown or save the plan to Excel to create a `Meal Prep` sheet.

## Workbook Settings and Saving

Planner settings are stored in the workbook's `Planner Settings` sheet; the built-in defaults are only used when a setting is missing. There is no in-app Settings tab. Weekly plans remain separate from the workbook. Select **Save plan to Excel** to create or replace the `Meal Plan`, `Shopping List`, and `Prep Notes` output sheets.

Workbook edits are written directly to the selected workbook. Automatic backups are disabled. The CLI writes a dated planned copy by default; `--inplace` updates the selected workbook directly:

```bash
python meal_planner.py
python meal_planner.py --inplace
python meal_planner.py --workbook path/to/food.xlsx
```