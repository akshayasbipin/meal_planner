"""Supabase access layer. All user-owned rows are additionally filtered by user_id."""
from __future__ import annotations
import json
import os
from dotenv import load_dotenv
from datetime import datetime, timezone
from io import BytesIO
import streamlit as st
from supabase import create_client

load_dotenv()

BUCKET = "user-workbooks"

def _client(url: str, key: str):
    # Auth state is mutable; create one client per Streamlit session, never globally cache it.
    return create_client(url, key)

def _setting(name: str):
    try:
        value = st.secrets[name]
    except Exception:
        value = None
    return value or os.getenv(name)

def get_client():
    url = _setting("SUPABASE_URL")
    key = _setting("SUPABASE_ANON_KEY")
    if not url or not key:
        raise RuntimeError("Add SUPABASE_URL and SUPABASE_ANON_KEY to .streamlit/secrets.toml or environment variables.")
    return _client(url, key)

def sign_up(client, email: str, password: str):
    return client.auth.sign_up({"email": email, "password": password})

def sign_in(client, email: str, password: str):
    return client.auth.sign_in_with_password({"email": email, "password": password})

def sign_out(client):
    client.auth.sign_out()

def upload_workbook(client, user_id: str, filename: str, content: bytes):
    safe_name = filename.rsplit("/", 1)[-1].replace("..", "_")
    path = f"{user_id}/{safe_name}"
    client.storage.from_(BUCKET).upload(path, content, {"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    payload = {"user_id": user_id, "storage_path": path, "file_name": safe_name,
               "updated_at": datetime.now(timezone.utc).isoformat()}
    client.table("user_workbooks").upsert(payload, on_conflict="user_id").execute()
    return path

def fetch_workbook(client, user_id: str):
    result = client.table("user_workbooks").select("storage_path,file_name,updated_at").eq("user_id", user_id).maybe_single().execute()
    if not result.data:
        return None
    meta = result.data
    data = client.storage.from_(BUCKET).download(meta["storage_path"])
    return meta, data

def save_rules(client, user_id: str, rules: list[dict]):
    client.table("user_rules").delete().eq("user_id", user_id).execute()
    rows = []
    for rule in rules:
        row = {"user_id": user_id}
        for field in ("rule_type", "day", "meal", "dish", "value"):
            value = rule.get(field)
            if isinstance(value, float) and value != value:  # NaN from an edited dataframe
                value = None
            row[field] = value
        row["active"] = bool(rule.get("active", True))
        rows.append(row)
    if rows:
        client.table("user_rules").insert(rows).execute()

def fetch_rules(client, user_id: str) -> list[dict]:
    result = client.table("user_rules").select("id,rule_type,day,meal,dish,value,active").eq("user_id", user_id).execute()
    return result.data or []

def save_plan(client, user_id: str, monday: str, days: list, plan: dict, notes: list):
    client.table("user_plans").upsert({"user_id": user_id, "week_start": monday, "days": days,
        "plan_data": plan, "notes": notes, "updated_at": datetime.now(timezone.utc).isoformat()},
        on_conflict="user_id,week_start").execute()

def fetch_plan(client, user_id: str, monday: str):
    result = client.table("user_plans").select("days,plan_data,notes").eq("user_id", user_id).eq("week_start", monday).maybe_single().execute()
    return result.data if result.data else None

def save_pantry(client, user_id: str, rows: list[dict]):
    client.table("user_pantry").delete().eq("user_id", user_id).execute()
    payload = [{"user_id": user_id, "item": r["item"], "quantity": r.get("qty"), "unit": r.get("unit", "")} for r in rows]
    if payload:
        client.table("user_pantry").insert(payload).execute()

def fetch_pantry(client, user_id: str):
    result = client.table("user_pantry").select("item,quantity,unit").eq("user_id", user_id).execute()
    return [{"item": r["item"], "qty": float(r["quantity"]) if r["quantity"] is not None else None, "unit": r["unit"] or ""} for r in (result.data or [])]

def save_nutrition_profile(client, user_id: str, entries: list[dict]):
    client.table("nutrition_profiles").delete().eq("user_id", user_id).execute()
    rows = [{"user_id": user_id, **e} for e in entries]
    if rows:
        client.table("nutrition_profiles").insert(rows).execute()

def fetch_nutrition_profile(client, user_id: str):
    result = client.table("nutrition_profiles").select("nutrient,reported_value,unit,reference_range,target,notes").eq("user_id", user_id).execute()
    return result.data or []
