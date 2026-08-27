# -*- coding: utf-8 -*-
"""
МойСклад → Google Sheets: "Заказы покупателей".

МУҲИМ ТЕХНИК ҚАРОР: HTTP сўровлар Python'нинг urllib'и билан эмас, СИСТЕМА
curl'и орқали юборилади. Сабаби — МойСклад'нинг nginx'и Python-urllib
сарлавҳаларини рад этади (HTTP 415). Шунингдек Accept сарлавҳаси АЙНАН
"application/json;charset=utf-8" бўлиши шарт (акс ҳолда HTTP 400).

ИШЛАШИ:
  1. Шитсдаги мавжуд қаторлар ўқилади (MS_ID бўйича харита қурилади)
  2. МойСклад'дан буюртмалар олинади:
       - шитс бўш бўлса         → START_DATE'дан бошлаб (биринчи тўлдириш)
       - шитсда маълумот бўлса  → фақат сўнгги REFRESH_DAYS кунлик
         (статус 5-10 кунда ўзгариши мумкин — ортиқча юк бўлмаслиги учун
          ундан эскиси қайта текширилмайди)
  3. ЯНГИ буюртмалар қўшилади (ҳар маҳсулот алоҳида қаторда)
  4. МАВЖУД буюртмаларнинг ўзгарувчан майдонлари (статус, склад, логистика
     ва ҳ.к.) янгиланади — эски ҳолда қотиб қолмаслиги учун

Керакли env (start.sh):
  MS_TOKEN, MS_SHEET_ID, MS_SA_JSON, MS_WORKSHEET, MS_START_DATE
  MS_REFRESH_DAYS (ихтиёрий, default 10)
"""
import json
import logging
import os
import subprocess
import time
import urllib.parse
from datetime import datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO,
                     format="%(asctime)s %(levelname)s: %(message)s")
log = logging.getLogger("ms")

MS_TOKEN = os.environ.get("MS_TOKEN", "")
SHEET_ID = os.environ.get("MS_SHEET_ID", "")
SA_JSON = os.environ.get("MS_SA_JSON", "/root/moysklad/service_account.json")
WORKSHEET = os.environ.get("MS_WORKSHEET", "MoySklad")
START_DATE = os.environ.get("MS_START_DATE", "2026-08-22")
REFRESH_DAYS = int(os.environ.get("MS_REFRESH_DAYS", "10"))

API = "https://api.moysklad.ru/api/remap/1.2"

# Устунлар тартиби (фойдаланувчи белгилаган)
HEADERS = ["№", "дата", "исм", "телефон", "товар", "кол", "сумма", "статус",
           "склад", "роп", "сотувчи", "источник", "регион", "адресс",
           "изменения", "логистика", "комментария", "MS_ID", "формула",
           "Креатив", "Таргетолог", "Ии профиль", "Форма", "ID"]

COL = {name: i for i, name in enumerate(HEADERS)}       # 0-индексли
MS_ID_COL_1 = COL["MS_ID"] + 1                           # gspread 1-индексли

# Вақт ўтиб ЎЗГАРИШИ мумкин бўлган устунлар — ҳар ишга туширишда текширилади
MUTABLE = ["статус", "склад", "роп", "сотувчи", "источник", "регион",
           "адресс", "изменения", "логистика", "комментария", "ID"]

# Ҳозирча МойСкладда манбаси йўқ — кейинчалик тўлдирилади
EMPTY_COLS = ["Креатив", "Таргетолог", "Ии профиль", "Форма"]


# ═══════════════════════ МойСклад API (curl орқали) ═══════════════════════

def ms_get(path, params=None, retries=3):
    """curl орқали GET сўров. Хатода None қайтаради."""
    url = API + path
    if params:
        url += "?" + urllib.parse.urlencode(params)

    cmd = [
        "curl", "-sS", "--compressed", "--max-time", "45",
        "-w", "\n__HTTP_CODE__%{http_code}",
        "-H", "Authorization: Bearer " + MS_TOKEN,
        "-H", "Accept: application/json;charset=utf-8",
        url,
    ]

    for attempt in range(1, retries + 1):
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            out = res.stdout
            if "__HTTP_CODE__" not in out:
                log.warning("curl javob bermadi (%d/%d): %s",
                            attempt, retries, res.stderr[:200])
                time.sleep(2)
                continue

            body, code = out.rsplit("__HTTP_CODE__", 1)
            code = code.strip()
            if code == "200":
                return json.loads(body)

            msg = body[:400]
            try:
                errs = json.loads(body).get("errors", [])
                msg = "; ".join(x.get("error", "") for x in errs)[:400]
            except Exception:
                pass
            log.error("MS HTTP %s (%s): %s", code, path, msg or "(bo'sh javob)")
            if code in ("401", "403"):
                return None
            time.sleep(2)

        except subprocess.TimeoutExpired:
            log.warning("curl timeout (%d/%d)", attempt, retries)
            time.sleep(2)
        except Exception as e:
            log.warning("curl xato (%d/%d): %s", attempt, retries, e)
            time.sleep(2)
    return None


def fetch_orders(since_date):
    """since_date'дан кейинги барча буюртмалар (пагинация билан)."""
    out, offset, limit = [], 0, 100
    while True:
        data = ms_get("/entity/customerorder", {
            "filter": f"moment>={since_date} 00:00:00",
            "order": "moment,asc",
            "limit": limit,
            "offset": offset,
            "expand": "agent,state,store,project,salesChannel",
        })
        if data is None:
            log.error("Buyurtmalarni olib bo'lmadi")
            return None
        rows = data.get("rows", [])
        if not rows:
            break
        out.extend(rows)
        total = data.get("meta", {}).get("size", 0)
        offset += limit
        log.info("   olindi: %d / %d", len(out), total)
        if offset >= total:
            break
        time.sleep(0.2)
    return out


def fetch_positions(order_id):
    """Буюртма маҳсулотлари (алоҳида сўров — фақат ЯНГИ буюртмалар учун)."""
    data = ms_get(f"/entity/customerorder/{order_id}/positions", {
        "limit": 100, "expand": "assortment",
    })
    return data.get("rows", []) if data else []


# ═══════════════════════ Ёрдамчилар ═══════════════════════

def attr(order, name):
    """Қўшимча майдон (доп. поле) қиймати — аниқ ном бўйича."""
    for a in (order.get("attributes") or []):
        if (a.get("name") or "").strip().lower() == name.lower():
            v = a.get("value")
            if isinstance(v, dict):
                return v.get("name") or v.get("value") or ""
            return str(v) if v is not None else ""
    return ""


def fmt_dt(raw, only_date=False):
    for f in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(raw, f)
            return dt.strftime("%d.%m.%Y") if only_date else dt.strftime("%d.%m.%Y %H:%M")
        except (ValueError, TypeError):
            continue
    return raw or ""


def day_key(raw):
    """Кунлик № ҳисоблаш учун — 'дд.мм.гггг'."""
    return fmt_dt(raw, only_date=True)


def order_fields(order):
    """Буюртмадан барча устун қийматларини (маҳсулотсиз) чиқаради."""
    agent = order.get("agent") or {}
    addr_full = order.get("shipmentAddressFull") or {}
    return {
        "дата":        fmt_dt(order.get("moment", "")),
        "исм":         agent.get("name", ""),
        "телефон":     (agent.get("phone") or "").strip(),
        "сумма":       int((order.get("sum") or 0) / 100),   # МойСклад тийинда
        "статус":      (order.get("state") or {}).get("name", ""),
        "склад":       (order.get("store") or {}).get("name", ""),
        "роп":         (order.get("project") or {}).get("name", ""),
        "сотувчи":     attr(order, "Продавцы (new)"),
        "источник":    (order.get("salesChannel") or {}).get("name", ""),
        "регион":      attr(order, "Регион"),
        "адресс":      order.get("shipmentAddress", "") or "",
        "изменения":   fmt_dt(order.get("updated", "")),
        "логистика":   attr(order, "Логистика"),
        "комментария": (addr_full.get("comment") or "").strip()
                        or (order.get("description") or "").strip(),
        "MS_ID":       order.get("id", ""),
        "формула":     1,
        "ID":          attr(order, "ID сделки в BX"),
    }


def build_rows(order, order_num):
    """Битта буюртма → Sheets қаторлари (ҳар маҳсулот алоҳида)."""
    f = order_fields(order)
    positions = fetch_positions(f["MS_ID"]) or [{}]

    rows = []
    for i, pos in enumerate(positions):
        tovar = (pos.get("assortment") or {}).get("name", "—")
        kol = pos.get("quantity", "")
        try:
            kol = int(float(kol))
        except (TypeError, ValueError):
            pass

        row = []
        for h in HEADERS:
            if h == "№":
                row.append(order_num if i == 0 else "")
            elif h == "товар":
                row.append(tovar)
            elif h == "кол":
                row.append(kol)
            elif h == "сумма":
                row.append(f["сумма"] if i == 0 else "")
            elif h in EMPTY_COLS:
                row.append("")
            else:
                row.append(f.get(h, ""))
        rows.append(row)
    return rows


# ═══════════════════════ Google Sheets ═══════════════════════

def get_ws():
    creds = Credentials.from_service_account_file(
        SA_JSON, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    book = gspread.authorize(creds).open_by_key(SHEET_ID)
    try:
        ws = book.worksheet(WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=WORKSHEET, rows=20000, cols=len(HEADERS) + 2)
        ws.append_row(HEADERS)
        return ws, []
    return ws, ws.get_all_values()


def col_letter(idx0):
    """0-индексдан A1 устун ҳарфига (0->A, 25->Z, 26->AA)."""
    n, s = idx0 + 1, ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    if not MS_TOKEN:
        raise SystemExit("❌ MS_TOKEN o'rnatilmagan")
    if not SHEET_ID:
        raise SystemExit("❌ MS_SHEET_ID o'rnatilmagan")

    ws, values = get_ws()

    # ── Сарлавҳа мослигини текширамиз ──
    if not values:
        ws.append_row(HEADERS)
        values = [HEADERS]
    elif values[0] != HEADERS:
        has_old_data = len(values) > 1 and any(r and any(r) for r in values[1:])
        if has_old_data:
            # ХАВФСИЗЛИК: сарлавҳани жимгина алмаштирсак, эски форматдаги
            # қаторлар устунлар бўйича СИЛЖИБ қолади (маълумот бузилади).
            log.error("❌ Шитсдаги сарлавҳа МОС ЭМАС, лекин эски маълумот бор!")
            log.error("   Кутилган: %s", " | ".join(HEADERS[:6]) + " ...")
            log.error("   Шитсда:   %s", " | ".join(values[0][:6]) + " ...")
            log.error("   Ечим: '%s' варағини БУТУНЛАЙ тозаланг (ёки ўчиринг), "
                      "кейин қайта ишга туширинг.", WORKSHEET)
            raise SystemExit(1)
        ws.update([HEADERS], f"A1:{col_letter(len(HEADERS)-1)}1")
        values[0] = HEADERS

    data_rows = values[1:]
    has_data = any(r and any(r) for r in data_rows)

    # ── Мавжуд буюртмалар харитаси: MS_ID -> [qator raqamlari] ──
    existing = {}
    day_max = {}          # 'дд.мм.гггг' -> энг катта №
    for i, r in enumerate(data_rows, start=2):   # 2 — биринчи маълумот қатори
        if len(r) <= COL["MS_ID"]:
            continue
        ms_id = r[COL["MS_ID"]]
        if ms_id:
            existing.setdefault(ms_id, []).append(i)
        num, dk = r[COL["№"]] if r else "", r[COL["дата"]][:10] if len(r) > COL["дата"] else ""
        if num and dk:
            try:
                day_max[dk] = max(day_max.get(dk, 0), int(num))
            except ValueError:
                pass

    log.info("Шитсда %d та буюртма (%d қатор)", len(existing), len(data_rows))

    # ── Қайси кундан бошлаб оламиз ──
    if has_data:
        since = max(START_DATE,
                    (datetime.now() - timedelta(days=REFRESH_DAYS)).strftime("%Y-%m-%d"))
        log.info("Режим: ЯНГИЛАШ (сўнгги %d кун, %s'дан)", REFRESH_DAYS, since)
    else:
        since = START_DATE
        log.info("Режим: БИРИНЧИ ТЎЛДИРИШ (%s'дан)", since)

    orders = fetch_orders(since)
    if orders is None:
        return
    log.info("МойСкладда %d та буюртма", len(orders))
    if not orders:
        return

    # ═══ 1) МАВЖУДЛАРНИ ЯНГИЛАШ (статус ва ҳ.к. ўзгарган бўлса) ═══
    updates, changed_ids = [], set()
    for o in orders:
        rows_idx = existing.get(o.get("id"))
        if not rows_idx:
            continue
        f = order_fields(o)
        for rn in rows_idx:
            cur = data_rows[rn - 2]
            for h in MUTABLE:
                c = COL[h]
                old = cur[c] if len(cur) > c else ""
                new = str(f.get(h, ""))
                if old != new:
                    updates.append({"range": f"{col_letter(c)}{rn}", "values": [[new]]})
                    changed_ids.add(o.get("id"))

    if updates:
        # Google API чекловидан ошмаслик учун бўлакларга бўламиз
        for i in range(0, len(updates), 200):
            ws.batch_update(updates[i:i+200], value_input_option="USER_ENTERED")
            time.sleep(0.5)
        log.info("🔄 Янгиланди: %d та буюртма (%d катак)", len(changed_ids), len(updates))
    else:
        log.info("🔄 Ўзгарган буюртма йўқ")

    # ═══ 2) ЯНГИЛАРНИ ҚЎШИШ ═══
    new_orders = [o for o in orders if o.get("id") not in existing]
    if not new_orders:
        log.info("➕ Янги буюртма йўқ")
        log.info("✅ Тайёр")
        return

    log.info("➕ Янги: %d та (маҳсулотлари олинмоқда...)", len(new_orders))
    new_orders.sort(key=lambda o: o.get("moment", ""))

    new_rows = []
    for i, o in enumerate(new_orders, 1):
        dk = day_key(o.get("moment", ""))
        day_max[dk] = day_max.get(dk, 0) + 1        # кунлик тартиб рақами
        new_rows.extend(build_rows(o, day_max[dk]))
        if i % 25 == 0:
            log.info("   ... %d/%d", i, len(new_orders))

    ws.append_rows(new_rows, value_input_option="USER_ENTERED")
    log.info("✅ %d та янги буюртма (%d қатор) қўшилди", len(new_orders), len(new_rows))


if __name__ == "__main__":
    main()
