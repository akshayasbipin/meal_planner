# Local weekly meal planner (Excel + Ollama + Streamlit)

Plans 7 days of lunch/dinner/snack from your `Grocery.xlsx` using a local Ollama model, builds the shopping list
from your Pantry, and lets you tweak everything in a browser UI. Everything runs on your Mac.

## Folder
```text
meal_planner/
├── Grocery.xlsx        # your workbook
├── app.py              # the interactive UI  (streamlit run app.py)
├── meal_planner.py     # the engine + command-line version
├── ui_logic.py         # helpers for the UI
├── config.json         # your rules (also editable in the UI's Settings tab)
├── requirements.txt
├── plans/              # created automatically: one JSON per week (history)
└── README.md
```

## Setup (once)
```bash
ollama pull qwen2.5:7b
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Daily use
```bash
streamlit run app.py          # opens http://localhost:8501
python meal_planner.py --inplace   # no-UI version, good for a Sunday cron job
```
Keep the Ollama app running. Close `Grocery.xlsx` in Excel before saving from the UI.

## The UI tabs
- **This week** - Generate, swap any dish from a dropdown, lock days you like and regenerate the rest, save to Excel.
- **Shopping** - live list that follows your swaps. Tick what you bought, enter the quantity, and it goes into Pantry.
  Download/copy the list to send to the family.
- **Dishes** - edit status (Known / Learning / Not tried), meal type, ease and ingredients; add dishes.
- **Pantry** - edit what's at home.
- **Settings** - model, breakfast, learning-dinner target, repeats, what you can't cook yet.

## How the rules work
- Lunch = dishes marked **Known** whose meal type includes Lunch (or is blank).
- Dinner = any dinner-capable dish; at least `min_learning_dinners` per week must be non-Known (shown with 🆕).
- Roti (and anything in `cannot_cook_yet`), breakfast-only dishes and Rice are never picked. Rice is added as a side.
- The model's answer is validated by code; near-miss spellings are fixed, rule breaks trigger a retry, and as a last
  resort Python swaps in valid dishes.
- The shopping list is computed by Python from the dishes' ingredients and your Pantry, not guessed by the model.

## Limits
This is a planning helper, not a nutritionist. Food supports but does not replace your clinician's treatment for the
vitamin D, folate and B12 findings.