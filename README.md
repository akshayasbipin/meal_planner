# Personalised Meal Planner

A Streamlit meal planner built around your Excel workbook, with Supabase authentication and per-user persistence. The existing rule-based meal-selection engine remains local and deterministic; no LLM is required to generate the weekly plan.

## What's included

- Supabase email/password sign-up and login.
- Workbook upload to a private Supabase Storage bucket, or restore the user's saved workbook.
- Per-user weekly plans, pantry, permanent rules and nutrition profile in Supabase Postgres.
- Permanent hard rules: fixed weekday meal, avoid dish, avoid ingredient.
- Soft preference: prefer ingredient (influences meal score).
- Nutrition Discovery: save a nutrient profile, research food sources and recipe ideas, and append a selected dish to the workbook.
- Grocery price research across Amazon Now, Flipkart Minutes, Instamart, Zepto and Blinkit using web-search results and retailer-specific search links.
- Search results are explicitly treated as leads. Price mentions extracted from snippets are not verified current offers.

## 1. Create Supabase project

1. Create a project at https://supabase.com.
2. Open **SQL Editor** and run `sql/supabase_schema.sql`.
3. In **Project Settings → API**, copy the Project URL and publishable/anon key.
4. Confirm email authentication is configured as you prefer. If email confirmation is enabled, users must confirm before signing in.

The SQL creates the private `user-workbooks` bucket, user-owned tables and Row Level Security policies. Do not put a Supabase `service_role` key in this app.

## 2. Configure secrets

For local development, create `.streamlit/secrets.toml` from `.streamlit/secrets.toml.example` and fill in the values. Do not commit this file.

`BRAVE_SEARCH_API_KEY` is optional. It enables in-app web research using Brave Search. Without it, the app still displays retailer-specific Google search links. Brave Search API pricing and terms can change; check its current dashboard before use.

## 3. Install and run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## 4. Workbook format

The uploaded `.xlsx` must contain these sheets (case-insensitive; spaces/underscores are interchangeable):

| Sheet | Purpose |
|---|---|
| `Grocery` | Grocery categories in row 1 and items below each category |
| `What can i make` | Dish, ingredients, ease, meal type, cooking status, optional link |
| `snacks` | Snack names in the first column |
| `Pantry` | Optional; created if missing |
| `Planner Settings` | Optional planner settings |
| `required` | Optional nutrient keyword priorities |

The workbook is saved back to the user's private Supabase Storage path when the app changes the pantry or adds a dish. Weekly plans and user rules are stored in separate Postgres tables.

## 5. Rule behaviour

- `fixed_meal`: hard rule. Example: Tuesday + Lunch + Curd Rice. It overrides a conflicting locked-day selection.
- `avoid_dish`: hard exclusion from generation.
- `avoid_ingredient`: hard exclusion if the ingredient appears in the dish name or ingredient text.
- `prefer_ingredient`: soft preference; increases the candidate's score but does not guarantee it.

A fixed dish must exist in the workbook and be eligible for the selected meal according to the workbook's meal type, familiarity and cooking settings.

## 6. Price research limitations

The research tab searches each named retailer's indexed public pages. Search engines may not index quick-commerce catalogue data, and search snippets can be stale or location-independent. The app does not bypass logins, bot protection or private APIs. Always open the retailer result and verify the exact product, pack size, seller, delivery fee, availability and delivery postcode before comparing. The initial version does not claim to provide real-time checkout prices.

## 7. Security and deployment notes

- RLS policies scope database rows to `auth.uid()`.
- Workbook files are in a private Storage bucket and paths start with the user's UUID.
- Only the Supabase publishable/anon key belongs in Streamlit secrets; never expose the service-role key.
- For a public deployment, configure secrets in the hosting provider rather than committing `.streamlit/secrets.toml`.
- Add rate limiting and abuse monitoring before opening sign-up publicly.

## 8. Current scope / next improvements

This first integrated version stores workbook, pantry, rules, plans and nutrition profile. It intentionally does not yet implement a full history browser, account deletion workflow, password reset UI, or a canonical nutrition database with nutrient quantities per serving. The web research layer can later be expanded with verified retailer feeds or authorised affiliate/catalogue APIs.
