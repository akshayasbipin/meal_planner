import unittest
from collections import Counter
from datetime import date
from types import SimpleNamespace

import meal_planner as mp
import ui_logic as ui
from openpyxl import Workbook


def make_context():
    lunch_pool = [
        {"name": f"Lunch {i}", "ingredients": "spinach lentil" if i == 0 else "paneer",
         "status": "known", "ease": "Easy"}
        for i in range(5)
    ]
    dinner_pool = [
        {"name": f"Dinner {i}", "ingredients": "methi dal" if i == 0 else "egg",
         "status": "known", "ease": "Easy"}
        for i in range(5)
    ]
    return SimpleNamespace(
        dishes=lunch_pool + dinner_pool,
        lunch_pool=lunch_pool,
        dinner_pool=dinner_pool,
        snacks=["Apple", "Nuts"],
        nutrition_priorities=[
            {"name": "Workbook priority A", "foods": ["spinach", "methi"], "target": 4},
            {"name": "Workbook priority B", "foods": ["egg", "paneer"], "target": 8},
        ],
        lim={"cap_l": 2, "cap_d": 2},
    )


class GenerateWeekTests(unittest.TestCase):
    def test_meal_prep_report_lists_missing_items_and_ingredient_gaps(self):
        ctx = make_context()
        days = [{"day": "Monday", "lunch": "Lunch 0", "dinner": "Dinner 0", "snack": "Apple"}]
        rows = [
            ["Spinach", "Vegetable", "Buy", "Lunch 0", 1, "", "", ""],
            ["Egg", "Proteins", "In stock", "Breakfast", 7, "12 pcs", "", ""],
            ["Unlisted spice", "Not in Grocery sheet", "Check / optional", "Dinner 0", 1, "", "", ""],
        ]
        text = mp.meal_prep_text(days, ctx.dishes, rows, dict(mp.DEFAULT_CONFIG), date(2026, 11, 16))
        buy_section = text.split("## Check manually", maxsplit=1)[0]

        self.assertIn("Spinach (needed for: Lunch 0)", buy_section)
        self.assertNotIn("Egg", buy_section)
        self.assertIn("Unlisted spice", text)
        self.assertIn("Ingredients:", text)

    def test_planner_settings_are_read_from_excel_with_defaults(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Planner Settings"
        ws.append(["Setting", "Value"])
        ws.append(["fixed_breakfast", "Oats"])
        ws.append(["max_repeats_lunch", 3])
        cfg = mp.load_config(wb)

        self.assertEqual(cfg["fixed_breakfast"], "Oats")
        self.assertEqual(cfg["max_repeats_lunch"], 3)
        self.assertEqual(cfg["weekly_staples"], mp.DEFAULT_CONFIG["weekly_staples"])

    def test_nutrient_priorities_are_read_generically_from_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "required"
        ws.append(["Nutrient", "Foods to include", "Weekly target servings"])
        ws.append(["Workbook nutrient", "spinach, lentils", 6])

        self.assertEqual(mp.read_nutrition_priorities(ws), [
            {"name": "Workbook nutrient", "foods": ["spinach", "lentils"], "target": 6}
        ])

    def test_manual_options_exclude_conflicts_but_keep_current_value(self):
        options = ui.meal_options(["A", "B", "C"], "Legacy", "A", {"C"})
        self.assertEqual(options, ["B", "Legacy"])

    def test_week_obeys_meal_and_repeat_rules(self):
        ctx = make_context()
        _, days, _ = mp.generate(ctx, dict(mp.DEFAULT_CONFIG))
        slots = [[day["lunch"], day["dinner"]] for day in days]

        self.assertEqual(len(days), 7)
        self.assertTrue(all(lunch != dinner for lunch, dinner in slots))
        self.assertTrue(all(set(slots[i]).isdisjoint(slots[i - 1]) for i in range(1, 7)))
        self.assertLessEqual(max(Counter(day["lunch"] for day in days).values()), 2)
        self.assertLessEqual(max(Counter(day["dinner"] for day in days).values()), 2)
        self.assertTrue(all(day["lunch"] in {d["name"] for d in ctx.lunch_pool} for day in days))
        self.assertTrue(all(day["dinner"] in {d["name"] for d in ctx.dinner_pool} for day in days))

    def test_workbook_priorities_are_preferred_and_locked_snack_is_preserved(self):
        ctx = make_context()
        _, days, _ = mp.generate(
            ctx,
            dict(mp.DEFAULT_CONFIG),
            {0: {"lunch": "Lunch 0", "dinner": "Dinner 1", "snack": "Pear"}},
        )

        green_days = sum("Lunch 0" == day["lunch"] or "Dinner 0" == day["dinner"] for day in days)
        self.assertGreaterEqual(green_days, 2)
        self.assertEqual(days[0]["snack"], "Pear")

    def test_impossible_locked_day_fails_with_actionable_error(self):
        ctx = make_context()
        with self.assertRaisesRegex(ValueError, "locked meals"):
            mp.generate(ctx, dict(mp.DEFAULT_CONFIG),
                        {0: {"lunch": "Lunch 0", "dinner": "Lunch 0"}})


if __name__ == "__main__":
    unittest.main()