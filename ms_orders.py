# -*- coding: utf-8 -*-
"""
МойСклад → Google Sheets: "Заказы покупателей" (мижоз буюртмалари).

Ишлаши:
  - МойСклад API'дан START_DATE'дан кейинги буюртмаларни олади
  - Ҳар маҳсулот АЛОҲИДА қаторга ёзилади (Сумма фақат биринчисига)
  - Аллақачон ёзилган буюртмалар (id бўйича) ҚАЙТА ёзилмайди
  - Cron орқали ҳар N дақиқада ишлатилади

Керакли env (start.sh'да):
  MS_TOKEN          — МойСклад API токени
  MS_SHEET_ID       — Google шитс ID
  MS_SA_JSON        — service_account.json йўли
  MS_WORKSHEET      — варақ номи (default: MoySklad)
  MS_START_DATE     — қайси кундан бошлаб (default: 2026-08-15)
"""
import json
import logging
import os
import time
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO,
                     format="%(asctime)s %(levelname)s: %(message)s")
log = logging.getLogger("ms")

# ═══════════════════════ Созламалар ═══════════════════════
MS_TOKEN = os.environ.get("MS_TOKEN", "")
SHEET_ID = os.environ.get("MS_SHEET_ID", "")
SA_JSON = os.environ.get("MS_SA_JSON", "/root/moysklad/service_account.json")
WORKSHEET = os.environ.get("MS_WORKSHEET", "MoySklad")
START_DATE = os.environ.get("MS_START_DATE", "2026-08-15")

API = "https://api.moysklad.ru/api/remap/1.2"

HEADERS = ["Sana", "Mijoz", "Telefon", "Mahsulot", "Soni", "Summa", "Status",
           "Sklad", "Proyekt_ROP", "Prodavcy", "Manba", "Region", "Manzil",
           "Ozgarish", "Logistika", "Izoh", "MS_ID"]
MS_ID_COL = len(HEADERS)  # охирги устун — такрорланмаслик учун

# Қўшимча майдонлар (доп. поля) — номи бўйича қидирилади
CUSTOM_FIELDS = {
    "Prodavcy": ["продавцы", "продавец", "prodavcy"],
    "Region": ["регион", "region"],
    "Logistika": ["логистика", "logistika"],
    "Proyekt_ROP": ["проект", "proyekt", "проект роп"],
    "Manba": ["канал продаж", "источник", "manba"],
    "Manzil": ["адрес доставки", "адрес", "manzil"],
}


# ═══════════════════════ МойСклад API ═══════════════════════

def ms_get(path, params=None, retries=3):
    url = API + path
    if params:
        parts = []
        for k, v in params.items():
            parts.append(f"{k}={urllib.parse.quote(str(v), safe='=;.,><~')}")
        url += "?" + "&".join(parts)

    for attempt in range(1, retries + 1):
        try:
            # МУҲИМ: МойСклад Python'нинг стандарт User-Agent'ини РАД ЭТАДИ
            # (HTTP 415). Шунинг учун curl'га ўхшаш сарлавҳалар юборилади.
            req = urllib.request.Request(url, headers={
                "Authorization": "Bearer " + MS_TOKEN,
                "Accept": "*/*",
                "User-Agent": "curl/8.5.0",
                "Accept-Encoding": "gzip",
            })
            with urllib.request.urlopen(req, timeout=40) as r:
                raw = r.read()
                if r.headers.get("Content-Encoding") == "gzip":
                    import gzip
                    raw = gzip.decompress(raw)
                return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="ignore")
            try:
                errs = json.loads(body).get("errors", [])
                msg = "; ".join(x.get("error", "") for x in errs)[:300]
            except Exception:
                msg = body[:300]
            log.error("MS HTTP %s (%s): %s", e.code, path, msg or "(bo'sh)")
            if e.code in (401, 403):
                return None  # токен хато — қайта уринишдан фойда йўқ
            time.sleep(2)
        except Exception as e:
            log.warning("MS urinish %d/%d: %s", attempt, retries, e)
            time.sleep(2)
    return None


def fetch_orders(since_date):
    """START_DATE'дан кейинги барча Заказы покупателей."""
    out = []
    offset = 0
    limit = 100
    while True:
        data = ms_get("/entity/customerorder", {
            "filter": f"moment>={since_date} 00:00:00",
            "order": "moment,asc",
            "limit": limit,
            "offset": offset,
            # МУҲИМ: filter билан бирга КЎП expand ишламайди (HTTP 400) —
            # шунинг учун фақат зарурларини сўраймиз, positions алоҳида олинади
            "expand": "agent,state,store,project",
        })
        if data is None:
            log.error("Buyurtmalarni olib bo'lmadi")
            break
        rows = data.get("rows", [])
        if not rows:
            break
        out.extend(rows)
        offset += limit
        if offset >= data.get("meta", {}).get("size", 0):
            break
        time.sleep(0.3)  # API'ни ортиқча юкламаслик
    return out


def fetch_positions(order_id):
    """Буюртма маҳсулотларини алоҳида сўров билан олади (expand чеклови сабабли)."""
    data = ms_get(f"/entity/customerorder/{order_id}/positions", {
        "limit": 100,
        "expand": "assortment",
    })
    if not data:
        return []
    return data.get("rows", [])


def get_attr(order, keys):
    """Доп. поля (attributes) ичидан ном бўйича қидиради."""
    for a in (order.get("attributes") or []):
        name = (a.get("name") or "").strip().lower()
        if any(k in name for k in keys):
            v = a.get("value")
            if isinstance(v, dict):
                return v.get("name") or v.get("value") or ""
            return str(v) if v is not None else ""
    return ""


def get_phone(agent):
    if not agent:
        return ""
    for k in ("phone", "actualAddress"):
        v = agent.get(k)
        if k == "phone" and v:
            return str(v).strip()
    return ""


def parse_order(order):
    """Битта буюртмани Sheets қаторларига айлантиради (ҳар маҳсулот — алоҳида)."""
    ms_id = order.get("id", "")
    moment = order.get("moment", "")
    try:
        dt = datetime.strptime(moment, "%Y-%m-%d %H:%M:%S.%f")
        sana = dt.strftime("%d.%m.%Y %H:%M")
    except (ValueError, TypeError):
        sana = moment

    agent = order.get("agent") or {}
    mijoz = agent.get("name", "")
    telefon = get_phone(agent)

    state = (order.get("state") or {}).get("name", "")
    store = (order.get("store") or {}).get("name", "")
    project = (order.get("project") or {}).get("name", "")
    izoh = order.get("description", "") or ""
    summa = (order.get("sum") or 0) / 100  # МойСклад тийинда сақлайди

    updated = order.get("updated", "")
    try:
        ud = datetime.strptime(updated, "%Y-%m-%d %H:%M:%S.%f")
        ozgarish = ud.strftime("%d.%m.%Y %H:%M")
    except (ValueError, TypeError):
        ozgarish = updated

    # Доп. поля
    proyekt_rop = project or get_attr(order, CUSTOM_FIELDS["Proyekt_ROP"])
    prodavcy = get_attr(order, CUSTOM_FIELDS["Prodavcy"])
    manba = get_attr(order, CUSTOM_FIELDS["Manba"])
    region = get_attr(order, CUSTOM_FIELDS["Region"])
    manzil = order.get("shipmentAddress", "") or get_attr(order, CUSTOM_FIELDS["Manzil"])
    logistika = get_attr(order, CUSTOM_FIELDS["Logistika"])

    positions = fetch_positions(ms_id)
    rows = []
    if not positions:
        positions = [{}]

    for i, pos in enumerate(positions):
        assortment = pos.get("assortment") or {}
        mahsulot = assortment.get("name", "—")
        soni = pos.get("quantity", "")
        try:
            soni = int(float(soni))
        except (TypeError, ValueError):
            pass

        rows.append([
            sana,                              # Sana
            mijoz,                             # Mijoz
            telefon,                           # Telefon
            mahsulot,                          # Mahsulot
            soni,                              # Soni
            int(summa) if i == 0 else "",      # Summa (фақат биринчи қаторга)
            state,                             # Status
            store,                             # Sklad
            proyekt_rop,                       # Proyekt_ROP
            prodavcy,                          # Prodavcy
            manba,                             # Manba
            region,                            # Region
            manzil,                            # Manzil
            ozgarish,                          # Ozgarish
            logistika,                         # Logistika
            izoh,                              # Izoh
            ms_id,                             # MS_ID (такрорланмаслик учун)
        ])
    return rows


# ═══════════════════════ Google Sheets ═══════════════════════

def get_ws():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(SA_JSON, scopes=scopes)
    book = gspread.authorize(creds).open_by_key(SHEET_ID)
    try:
        return book.worksheet(WORKSHEET)
    except Exception:
        ws = book.add_worksheet(title=WORKSHEET, rows=10000, cols=len(HEADERS) + 2)
        ws.append_row(HEADERS)
        return ws


def existing_ids(ws):
    """Аллақачон ёзилган буюртма ID'лари (такрор ёзмаслик учун)."""
    try:
        col = ws.col_values(MS_ID_COL)
        return set(col[1:])  # header'ни ташлаб
    except Exception as e:
        log.error("existing_ids: %s", e)
        return set()


def main():
    if not MS_TOKEN:
        raise SystemExit("❌ MS_TOKEN o'rnatilmagan (start.sh'ni tekshiring)")
    if not SHEET_ID:
        raise SystemExit("❌ MS_SHEET_ID o'rnatilmagan")

    log.info("МойСклад'дан буюртмалар олинмоқда (%s'дан)...", START_DATE)
    orders = fetch_orders(START_DATE)
    log.info("Жами %d та буюртма топилди", len(orders))
    if not orders:
        return

    ws = get_ws()
    already = existing_ids(ws)
    log.info("Шитсда аллақачон %d та буюртма бор", len(already))

    new_orders = [o for o in orders if o.get("id") not in already]
    log.info("Янги буюртмалар: %d та (ҳар бири учун маҳсулотлар олинмоқда...)", len(new_orders))

    new_rows = []
    new_count = 0
    for i, order in enumerate(new_orders, 1):
        new_rows.extend(parse_order(order))
        new_count += 1
        if i % 25 == 0:
            log.info("   ... %d/%d", i, len(new_orders))

    if not new_rows:
        log.info("Янги буюртма йўқ")
        return

    ws.append_rows(new_rows, value_input_option="USER_ENTERED")
    log.info("✅ %d та янги буюртма (%d қатор) ёзилди", new_count, len(new_rows))


if __name__ == "__main__":
    main()