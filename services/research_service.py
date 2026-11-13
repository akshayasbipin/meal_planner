"""Search-based research. Brave results are leads, not confirmed live offers."""
from __future__ import annotations
import os, re
import requests
from dotenv import load_dotenv

load_dotenv()

PLATFORMS = {
    "Amazon Now": "amazon.in",
    "Flipkart Minutes": "flipkart.com",
    "Instamart": "swiggy.com/instamart",
    "Zepto": "zepto.com",
    "Blinkit": "blinkit.com",
}

def search_web(query: str, count: int = 8) -> list[dict]:
    key = os.getenv("BRAVE_SEARCH_API_KEY")
    if not key:
        raise RuntimeError("Set BRAVE_SEARCH_API_KEY to enable in-app web research. Search links are available without it.")
    response = requests.get("https://api.search.brave.com/res/v1/web/search",
        params={"q": query, "count": count, "country": "IN", "search_lang": "en", "safesearch": "moderate"},
        headers={"Accept": "application/json", "X-Subscription-Token": key}, timeout=20)
    response.raise_for_status()
    data = response.json()
    return [{"title": r.get("title", ""), "url": r.get("url", ""),
             "description": r.get("description", ""), "source": (r.get("profile") or {}).get("long_name", "Web search"),
             "price_mentions": re.findall(r"(?:₹|Rs\.?\s?)(\d[\d,]*(?:\.\d{1,2})?)", (r.get("title", "") + " " + r.get("description", "")))}
            for r in data.get("web", {}).get("results", [])]

def search_grocery_item(item: str, area: str = "Hyderabad") -> dict:
    if not os.getenv("BRAVE_SEARCH_API_KEY"):
        raise RuntimeError("Set BRAVE_SEARCH_API_KEY to enable in-app web research. Retailer-specific search links are shown below without it.")
    results = {}
    for platform, domain in PLATFORMS.items():
        query = f'{item} price {area} site:{domain}'
        try:
            results[platform] = search_web(query, count=4)
        except Exception as exc:
            results[platform] = {"error": str(exc)}
    return results

def search_nutrition(nutrient: str) -> list[dict]:
    return search_web(f"Indian foods rich in {nutrient} nutrition database ICMR NIH", count=8)

def search_recipes(food: str) -> list[dict]:
    return search_web(f"healthy Indian recipes containing {food}", count=8)
