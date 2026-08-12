import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import csv, math, io, re
from datetime import date, timedelta
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import supplier_allocation as supplier_logic

allocate_orders = supplier_logic.allocate_orders
build_reference_price_map = supplier_logic.build_reference_price_map
google_sheet_csv_url = supplier_logic.google_sheet_csv_url
google_sheet_xlsx_url = getattr(
    supplier_logic,
    "google_sheet_xlsx_url",
    lambda url: re.sub(r"/export.*$", "/export?format=xlsx", google_sheet_csv_url(url)),
)
_supplier_price_parser = supplier_logic.parse_supplier_prices

# Compatibility with an older supplier_allocation.py that may still be cached on
# Streamlit Cloud. The current module provides these functions directly; the local
# versions keep app.py importable while the second file is being refreshed.
def _normalise_product_code(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        return str(int(value)) if value.is_integer() else format(value, "f").rstrip("0").rstrip(".")
    text = str(value).replace("\xa0", " ").strip().strip("'\"")
    text = re.sub(r"\s+", "", text)
    return text[:-2] if re.fullmatch(r"[+-]?\d+\.0", text) else text


def _fallback_parse_barcode_mapping_rows(rows):
    values = [tuple(row) for row in rows]
    article_col = barcode_col = header_index = None
    for row_index, row in enumerate(values[:10]):
        for col_index, value in enumerate(row):
            header = re.sub(r"\s+", " ", str(value or "").strip().lower().replace("_", " "))
            if header in ("артикул", "sku", "код товару", "код товара", "article"):
                article_col = col_index
            if header in ("штрихкод", "штрих-код", "штрих код", "barcode", "ean", "upc"):
                barcode_col = col_index
        if article_col is not None and barcode_col is not None:
            header_index = row_index
            break
    if header_index is None:
        return {}, {"error": "Не знайдено колонки «Артикул» і «Штрихкод»"}
    mapping = {}; conflicts = []; article_barcodes = {}; empty_rows = 0; valid_pairs = 0
    for row in values[header_index + 1:]:
        article = str(row[article_col] or "").strip() if article_col < len(row) else ""
        barcode = _normalise_product_code(row[barcode_col] if barcode_col < len(row) else None)
        if not article or not barcode:
            empty_rows += 1; continue
        if barcode in mapping and mapping[barcode] != article:
            conflicts.append({"barcode": barcode, "articles": sorted({mapping[barcode], article})})
            continue
        mapping[barcode] = article
        article_barcodes.setdefault(article, set()).add(barcode)
        valid_pairs += 1
    return mapping, {
        "total_rows": max(0, len(values) - header_index - 1), "valid_pairs": valid_pairs,
        "unique_barcodes": len(mapping), "unique_articles": len(article_barcodes),
        "articles_with_multiple_barcodes": sum(len(codes) > 1 for codes in article_barcodes.values()),
        "empty_rows": empty_rows, "conflicts": conflicts, "error": "",
    }


def _fallback_apply_barcode_mapping(offers, barcode_map):
    remapped = {}; mapped = 0; direct = 0; unknown = []; duplicates = []
    for source_code, source_offer in offers.items():
        code = _normalise_product_code(source_code)
        article = barcode_map.get(code)
        if article:
            mapped += 1
        elif re.search(r"[A-Za-zА-Яа-яІіЇїЄє]", code):
            article = str(source_code).strip(); direct += 1
        else:
            unknown.append(code); continue
        offer = dict(source_offer, sku=article, supplier_sku=str(source_code).strip(),
                     barcode=code if code in barcode_map else "")
        if article in remapped:
            duplicates.append({"article": article, "supplier_codes": [remapped[article].get("supplier_sku", ""), offer["supplier_sku"]]})
            offer = min([remapped[article], offer], key=lambda item: (0 if item.get("in_stock") else 1, item.get("price", math.inf), -(item.get("available_qty") or 0)))
        remapped[article] = offer
    return remapped, {"source_offers": len(offers), "mapped_by_barcode": mapped, "direct_sku": direct,
                      "unknown_codes": sorted(set(filter(None, unknown))), "duplicate_articles": duplicates}


parse_barcode_mapping_rows = getattr(
    supplier_logic, "parse_barcode_mapping_rows", _fallback_parse_barcode_mapping_rows)
apply_barcode_mapping = getattr(
    supplier_logic, "apply_barcode_mapping", _fallback_apply_barcode_mapping)

# Prevent the empty-candidate crash present in one earlier deployed version.
if hasattr(supplier_logic, "_rank_candidates"):
    _original_rank_candidates = supplier_logic._rank_candidates
    def _safe_rank_candidates(candidates, rules):
        return _original_rank_candidates(candidates, rules) if candidates else []
    supplier_logic._rank_candidates = _safe_rank_candidates


def parse_supplier_prices(workbook, supplier_name, sheet_name=None):
    """Prefer a barcode column even when the deployed parser is an older version."""
    sheet = workbook[sheet_name] if sheet_name else workbook[workbook.sheetnames[0]]
    for row in sheet.iter_rows(min_row=1, max_row=min(10, sheet.max_row)):
        barcode_cell = None
        sku_cells = []
        for cell in row:
            header = re.sub(r"\s+", " ", str(cell.value or "").strip().lower().replace("_", " "))
            if header in ("штрихкод", "штрих-код", "штрих код", "barcode", "ean", "upc"):
                barcode_cell = cell
            elif header in ("артикул", "sku", "код товару", "код товара", "item code", "product code"):
                sku_cells.append(cell)
        if barcode_cell:
            for cell in sku_cells:
                cell.value = "Код постачальника"
            barcode_cell.value = "SKU"
            break
    try:
        return _supplier_price_parser(workbook, supplier_name, sheet_name=sheet.title)
    except TypeError:
        # Older deployed parser accepts only the first tab. Copy the selected tab
        # to a temporary workbook so it can still be scanned.
        if sheet.title == workbook.sheetnames[0]:
            return _supplier_price_parser(workbook, supplier_name)
        temporary = Workbook(); target = temporary.active; target.title = sheet.title[:31]
        for row in sheet.iter_rows(values_only=True):
            target.append(list(row))
        return _supplier_price_parser(temporary, supplier_name)

st.set_page_config(
    page_title="ProcureAI — Аналіз закупівель",
    page_icon="📦", layout="wide",
    initial_sidebar_state="expanded",
)

UA_MO = {'Січень':1,'Лютий':2,'Березень':3,'Квітень':4,'Травень':5,'Червень':6,
         'Липень':7,'Серпень':8,'Вересень':9,'Жовтень':10,'Листопад':11,'Грудень':12}
BUILT_IN_K = [1.1,1.15,1.18,1.05,0.95,1.34,1.34,0.9,0.92,1.2,1.5,1.56]

# ── Хелпери ──────────────────────────────────
def cs(v):
    """clean string — прибирає \xa0 та зайві пробіли"""
    return str(v).replace('\xa0', ' ').strip() if v is not None else ''

def sn(v):
    """safe number — конвертує рядок з пробілами і комами"""
    if v is None: return None
    try:
        f = float(cs(v).replace(' ', '').replace(',', '.'))
        return f
    except: return None

def parse_disc(v):
    """'Знижка 15 %', '15%', 0.15, 15 → 15.0"""
    if v is None: return 0
    s = cs(v)
    if not s or s in ('-', ''): return 0
    nums = re.findall(r'\d+[.,]?\d*', s)
    if not nums: return 0
    try:
        val = float(nums[0].replace(',', '.'))
        if val < 1: val *= 100
        return round(val, 1) if val > 0 else 0
    except: return 0

def mo_num(label):
    return UA_MO.get(cs(label).split()[0], 0)

def mo_year(label):
    for p in cs(label).split():
        if p.isdigit() and len(p) == 4: return int(p)
    return 0

# ── Завантаження Excel ────────────────────────
@st.cache_data(show_spinner=False)
def load_wb(file_bytes):
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)


@st.cache_data(show_spinner=False)
def load_barcode_mapping():
    """Load the repository copy converted from the user's legacy XLS table."""
    path = Path(__file__).with_name("barcode_mapping.csv")
    if not path.exists():
        return {}, {"error": "Файл barcode_mapping.csv відсутній"}
    with path.open("r", encoding="utf-8-sig", newline="") as source:
        return parse_barcode_mapping_rows(csv.reader(source))


@st.cache_data(show_spinner=False, ttl=300)
def load_google_sheet_prices(sheet_url, supplier_name):
    """Read every tab of a public Google Sheet and select the actual price tab."""
    export_url = google_sheet_xlsx_url(sheet_url)
    request = Request(export_url, headers={"User-Agent": "ProcureAI/1.0"})
    try:
        with urlopen(request, timeout=20) as response:
            payload = response.read()
            content_type = response.headers.get("Content-Type", "")
    except HTTPError as exc:
        raise ValueError(f"Google Sheets повернув помилку HTTP {exc.code}") from exc
    except URLError as exc:
        raise ValueError("Не вдалося підключитися до Google Sheets") from exc

    text_preview = payload[:500].decode("utf-8-sig", errors="replace")
    if "text/html" in content_type.lower() or text_preview.lstrip().lower().startswith("<!doctype html"):
        raise ValueError("Таблиця не опублікована для перегляду за посиланням")

    if payload[:2] == b"PK":
        try:
            workbook = openpyxl.load_workbook(io.BytesIO(payload), data_only=True)
        except Exception as exc:
            raise ValueError("Google Sheets повернув пошкоджений Excel-файл") from exc
        candidates = []
        for sheet_name in workbook.sheetnames:
            offers, warnings = parse_supplier_prices(workbook, supplier_name, sheet_name=sheet_name)
            recognised = not any("заголовки не розпізнано" in warning for warning in warnings)
            candidates.append((recognised, len(offers), sheet_name, offers, warnings))
        if not candidates:
            raise ValueError("Google Sheet не містить вкладок")
        _, _, selected_sheet, offers, warnings = max(candidates, key=lambda item: (item[0], item[1]))
        if len(workbook.sheetnames) > 1:
            warnings = [f"{supplier_name}: обрано вкладку «{selected_sheet}»"] + warnings
        if not offers:
            sheet = workbook[selected_sheet]
            preview_headers = []
            for row in sheet.iter_rows(min_row=1, max_row=min(20, sheet.max_row), values_only=True):
                filled = [str(value).strip() for value in row if value not in (None, "")]
                if filled:
                    preview_headers.append(" | ".join(filled[:8]))
                if len(preview_headers) >= 3:
                    break
            warnings.append(
                f"{supplier_name}: на вкладці «{selected_sheet}» не знайдено товарів. "
                f"Перші рядки: {' / '.join(preview_headers) or 'порожньо'}"
            )
        return offers, warnings

    # Compatibility fallback for servers or tests returning the selected tab as CSV.
    text = payload.decode("utf-8-sig", errors="replace")
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        raise ValueError("Google Sheet не містить даних")
    workbook = Workbook(); worksheet = workbook.active
    for row in rows:
        worksheet.append(row)
    offers, warnings = parse_supplier_prices(workbook, supplier_name)
    if not offers:
        headers = " / ".join(" | ".join(cell for cell in row[:8] if cell) for row in rows[:3])
        warnings.append(f"{supplier_name}: не знайдено товарів. Перші рядки: {headers or 'порожньо'}")
    return offers, warnings

# ── Парсинг шаблону ───────────────────────────
def parse_template(wb):
    """
    Шаблон має 3 вкладки:
      'продажі дані'      — блоки по 5 колонок: Назва|Артикул|Кількість|Рентаб|_
      'наявність на складі' — Номенклатура|Артикул|Залишок|Замовлено
      'Залишки'           — блоки по 4 колонки: Артикул|Назва|Дні|_ (по місяцях)
    """
    errors = []

    # Знаходимо вкладки (нечутливо до регістру)
    sheets = {n.lower(): n for n in wb.sheetnames}
    def find(key):
        for k, v in sheets.items():
            if key in k: return wb[v]
        return None

    ws_sales = find('продаж')
    ws_stock = find('наявн')
    ws_avail = find('залиш')

    if ws_sales is None: errors.append("Не знайдено вкладку 'продажі дані'")
    if ws_stock is None: errors.append("Не знайдено вкладку 'наявність на складі'")
    if ws_avail is None: errors.append("Не знайдено вкладку 'Залишки'")
    if errors: return None, None, None, None, errors

    # ── Продажі ──
    all_rows = list(ws_sales.iter_rows(values_only=True))
    # Рядок 2 (idx 1) містить місяці "Період: Місяць YYYY р."
    row2 = all_rows[1] if len(all_rows) > 1 else []
    month_blocks = []
    for i, v in enumerate(row2):
        label = cs(v)
        if 'Період' in label and mo_num(label.replace('Період:', '').replace('Період :', '').strip().rstrip('.')) > 0:
            clean_label = label.replace('Період:', '').replace('Період :', '').strip().rstrip('.').strip()
            month_blocks.append((i, clean_label))  # 0-indexed col offset

    if not month_blocks:
        errors.append("Не знайдено місяців у вкладці 'продажі дані'")
        return None, None, None, None, errors

    months_labels = [lbl for _, lbl in month_blocks]

    # Дані з рядка 14 (idx 13)
    sku_data = {}
    DATA_START = 13
    for row in all_rows[DATA_START:]:
        for ci, label in month_blocks:
            sku_val  = row[ci+1] if ci+1 < len(row) else None
            name_val = row[ci]   if ci   < len(row) else None
            qty_val  = row[ci+2] if ci+2 < len(row) else None
            rent_val = row[ci+3] if ci+3 < len(row) else None

            if not sku_val: continue
            sku = cs(sku_val)
            if not sku or sku in ('Артикул', 'IHERB', '1*', ''): continue

            name = cs(name_val) if name_val else sku
            qty  = max(sn(qty_val) or 0, 0)
            rent = sn(rent_val)

            if sku not in sku_data:
                sku_data[sku] = {'name': name, 'months': {}}
            if label not in sku_data[sku]['months']:
                sku_data[sku]['months'][label] = [0.0, None]
            sku_data[sku]['months'][label][0] += qty
            if rent is not None and qty > 0:
                sku_data[sku]['months'][label][1] = rent

    # ── Наявність (залишок + в дорозі) ──
    stock_map = {}
    all_stock = list(ws_stock.iter_rows(values_only=True))
    # Рядок 2: заголовки, Рядок 4+: дані
    for row in all_stock[3:]:
        if len(row) < 2 or not row[1]: continue
        sku = cs(row[1])
        if not sku or sku == '1*': continue
        nm  = cs(row[0]) if row[0] else sku
        st  = max(sn(row[2]) or 0, 0)
        tr  = max(sn(row[3]) or 0, 0) if len(row) > 3 and row[3] else 0
        stock_map[sku] = (st, tr, nm)

    # ── Залишки (дні на складі по місяцях) ──
    avail_map = {}
    all_av = list(ws_avail.iter_rows(values_only=True))
    # Рядок 2 (idx 1): місяці, кожні 4 колонки: Артикул|Назва|Дні|_
    row2_av = all_av[1] if len(all_av) > 1 else []
    av_blocks = []
    for i, v in enumerate(row2_av):
        label = cs(v)
        if 'Період' in label:
            clean = label.replace('Період:', '').replace('Період :', '').strip().rstrip('.').strip()
            av_blocks.append((i, clean))

    for row in all_av[3:]:  # з рядка 4
        for ci, label in av_blocks:
            sku_val  = row[ci]   if ci   < len(row) else None
            days_val = row[ci+2] if ci+2 < len(row) else None
            if not sku_val: continue
            sku = cs(sku_val)
            if not sku or sku in ('Артикул', '1*', ''): continue
            days = max(sn(days_val) or 0, 0)
            if sku not in avail_map: avail_map[sku] = {}
            avail_map[sku][label] = avail_map[sku].get(label, 0) + days

    return sku_data, months_labels, stock_map, avail_map, []

# ── Парсинг файлу цін ─────────────────────────
def parse_prices(wb, supplier_name="iHerb"):
    """Compatibility wrapper around the configurable supplier price parser."""
    return parse_supplier_prices(wb, supplier_name)

# ── Основний аналіз ───────────────────────────
def run_analysis(sku_data, months_labels, stock_map, avail_map, price_map, params):
    today     = date.today()
    CUR_MO    = today.month; CUR_YR = today.year; CUR_DAY = today.day
    cur_scale = 30.0 / CUR_DAY

    MG_MIN    = params['mg_min']; LEAD = params['lead']
    SAFETY    = params['safety']; SAFETY60 = params['safety60']
    SAFETY90  = params.get('safety90', 90)
    A_MULT    = params['a_mult']; MG_A = params['mg_a']
    MIN_MO    = params['min_months']; MIN_QTY = params['min_qty']
    AA        = params['avail_alpha']; LAM = params['lambda_val']
    LOW_AV    = params['low_avail']
    DISC_THR  = params['disc_thr']; DISC_THR2 = params.get('disc_thr2', 40)

    n = len(months_labels)

    def is_cur(lbl):  return mo_num(lbl) == CUR_MO and mo_year(lbl) == CUR_YR
    def is_done(lbl):
        mn, yr = mo_num(lbl), mo_year(lbl)
        return yr < CUR_YR or (yr == CUR_YR and mn < CUR_MO)

    mo_complete = [is_done(lbl) for lbl in months_labels]
    mo_is_cur   = [is_cur(lbl)  for lbl in months_labels]
    exp_w = [math.exp(LAM * (i - (n-1))) for i in range(n)]

    # Pass 1 — season K
    mo_cs = [0.0] * n; mo_cc = [0] * n
    for sku, sd in sku_data.items():
        for i, lbl in enumerate(months_labels):
            qty, rent = sd['months'].get(lbl, [0.0, None])
            inc = (rent >= MG_MIN) if rent is not None else (qty > 0)
            if inc and qty > 0 and mo_complete[i]:
                mo_cs[i] += qty; mo_cc[i] += 1

    ca_  = [mo_cs[i]/mo_cc[i] if mo_cc[i] > 0 else 0 for i in range(n)]
    cv   = [ca_[i] for i in range(n) if mo_complete[i] and ca_[i] > 0]
    global_avg = sum(cv)/len(cv) if cv else 1.0
    cur_ks = [ca_[i]/global_avg for i, lbl in enumerate(months_labels)
              if mo_complete[i] and mo_num(lbl) == CUR_MO and ca_[i] > 0]
    season_K = sum(cur_ks)/len(cur_ks) if cur_ks else BUILT_IN_K[CUR_MO-1]

    # Pass 2 — per SKU
    results = []; excl_mg = []; sporadic = []

    for sku, sd in sku_data.items():
        st_info = stock_map.get(sku)
        stock   = st_info[0] if st_info else 0
        transit = st_info[1] if st_info else 0
        nm      = st_info[2] if st_info and st_info[2] else sd['name']
        eff     = stock + transit
        av_mo   = avail_map.get(sku, {})

        # Зважений avg/день
        ws_ = 0.0; wd_ = 0.0; cd_clean = 0.0; rd = n * 30
        for i, lbl in enumerate(months_labels):
            qty, rent = sd['months'].get(lbl, [0.0, None])
            inc = (rent >= MG_MIN) if rent is not None else (qty > 0)
            if not inc: continue
            ad = av_mo.get(lbl, 30)
            wt = exp_w[i]
            q_ = qty * (cur_scale if mo_is_cur[i] else 1)
            a_ = min(ad  * (cur_scale if mo_is_cur[i] else 1), 30)
            if a_ > 0:
                ws_ += q_ * wt; wd_ += a_ * wt
            cd_clean += a_

        avg_day   = ws_ / wd_ if wd_ > 0 else 0
        avail_pct = round(cd_clean / rd * 100) if rd > 0 else 0
        avail_K   = math.pow(avail_pct/100, AA) if avail_pct > 0 else 0

        # Середня маржа (тільки чисті місяці)
        rents = [r for lbl in months_labels
                 for q, r in [sd['months'].get(lbl, [0, None])]
                 if r is not None and r >= MG_MIN]
        avg_margin = round(sum(rents)/len(rents), 1) if rents else None

        # Фільтр маржі
        if avg_margin is not None and avg_margin < MG_MIN:
            cmo = sum(1 for lbl in months_labels if sd['months'].get(lbl,[0,None])[0]>0)
            ts  = sum(sd['months'].get(lbl,[0,None])[0] for lbl in months_labels)
            excl_mg.append({'sku':sku,'name':nm,'avg_margin':avg_margin,
                             'clean_months':cmo,'total_sold':round(ts,1)})
            continue

        # Постійний попит
        cmo = sum(1 for lbl in months_labels if sd['months'].get(lbl,[0,None])[0]>0)
        ts  = sum(sd['months'].get(lbl,[0,None])[0] for lbl in months_labels)
        is_sp = not (cmo >= MIN_MO and ts >= MIN_QTY)
        reason = ''
        if is_sp:
            parts = []
            if cmo < MIN_MO:  parts.append(f"{cmo}<{MIN_MO}міс")
            if ts  < MIN_QTY: parts.append(f"{round(ts):.0f}<{MIN_QTY}шт")
            reason = ', '.join(parts)

        # ABC
        if avg_margin is None:    abc, am = '?', 1.0
        elif avg_margin >= MG_A:  abc, am = 'A', A_MULT
        else:                     abc, am = 'B', 1.0

        # Тренд
        clean_m = [(lbl, q) for lbl in months_labels
                   for q, r in [sd['months'].get(lbl,[0,None])]
                   if q > 0 and av_mo.get(lbl, 30) > 0]
        if len(clean_m) >= 6:
            f3 = sum(q for _,q in clean_m[:3])/3
            l3 = sum(q for _,q in clean_m[-3:])/3
            tr_r = round(l3/f3, 2) if f3 > 0 else None
            trend = ('↑ зростає' if tr_r and tr_r > 1.3 else
                     '↓ спадає'  if tr_r and tr_r < 0.7 else '→ стабільно')
        else:
            l3 = sum(q for _,q in clean_m[-3:])/3 if clean_m else 0
            trend = '↑ новий' if l3 > 0 else '—'

        # Ціни
        pi = price_map.get(sku)
        p_avail = pi[1] if pi else False
        p_disc  = pi[2] if pi else 0
        p_price = pi[0] if pi else None
        use_60  = not is_sp and p_avail and p_disc >= DISC_THR
        use_90  = not is_sp and p_avail and p_disc >= DISC_THR2
        safety_disc = SAFETY90 if use_90 else (SAFETY60 if use_60 else 0)

        # Маржинальний дохід
        mi_day = None; sell_price = None
        if p_price and avg_margin and 0 < avg_margin < 100:
            sell_price = p_price / (1 - avg_margin/100)
            mi_day = round(avg_day * (sell_price - p_price), 4)

        dl = round(stock/avg_day) if avg_day > 0 else 999
        st2 = ('Критично' if dl<5 else 'Низько' if dl<15
               else 'Надлишок' if dl>90 else 'Норма')
        # Потреба залежить від попиту, а не від наявності в конкретного постачальника.
        # Відсутні пропозиції потраплять у список нерозподілених SKU після алокації.
        rec   = max(0,round(avg_day*avail_K*am*(LEAD+SAFETY  )*season_K-eff)) if not is_sp else 0
        rec60 = max(0,round(avg_day*avail_K*am*(LEAD+safety_disc)*season_K-eff)) if safety_disc else 0
        zero_date = (today+timedelta(days=int(dl))).strftime('%d.%m.%Y') if 0<=dl<999 else None

        row_d = dict(
            sku=sku, name=nm, abc=abc, avg_margin=avg_margin,
            avail_pct=avail_pct, avail_K=round(avail_K,3),
            trend=trend, avg_day=round(avg_day,4),
            stock=stock, transit=transit, eff_stock=eff,
            days_left=dl, zero_date=zero_date, status=st2,
            is_sporadic=is_sp, sporadic_reason=reason,
            season_K=round(season_K,3), rec=rec, rec_60=rec60,
            use_60=use_60, use_90=use_90, safety_disc=safety_disc,
            price_disc=round(p_disc,1), buy_price=p_price, sell_price=sell_price,
            mi_day=mi_day, low_avail=avail_pct < LOW_AV,
        )
        if is_sp: sporadic.append(row_d)
        else:     results.append(row_d)

    # ABC по MI (Pareto 70/90)
    mi_vals = sorted([r['mi_day'] for r in results if r['mi_day']], reverse=True)
    total_mi = sum(mi_vals); cum = 0; ta = tb = None
    for v in mi_vals:
        cum += v
        if ta is None and cum >= total_mi*0.70: ta = v
        if tb is None and cum >= total_mi*0.90: tb = v
    for r in results:
        mi = r['mi_day']
        if mi is None:          r['abc_mi'] = '?'
        elif ta and mi >= ta:   r['abc_mi'] = 'A'
        elif tb and mi >= tb:   r['abc_mi'] = 'B'
        else:                   r['abc_mi'] = 'C'

    meta = dict(season_K=season_K, global_avg=global_avg, n_months=n,
                cur_day=CUR_DAY, cur_scale=cur_scale,
                total_mi=total_mi, ta=ta, tb=tb, months=months_labels)
    return dict(regular=results, sporadic=sporadic, excl_mg=excl_mg, meta=meta)

# ── Генерація Excel ───────────────────────────
def gen_excel(data, params):
    results = data['regular']; sporadic = data['sporadic']
    meta    = data['meta']; today = date.today()

    th = Side(style="thin", color="BFBFBF")
    def tb(): return Border(left=th,right=th,top=th,bottom=th)
    def fl(c): return PatternFill("solid",start_color=c,fgColor=c)
    def hf(color="FFFFFF",bold=True,sz=10): return Font(name="Arial",bold=bold,color=color,size=sz)
    def cf(bold=False,sz=10,color="000000"): return Font(name="Arial",bold=bold,size=sz,color=color)
    def ca(): return Alignment(horizontal="center",vertical="center",wrap_text=True)
    def la(): return Alignment(horizontal="left",vertical="center")

    C = dict(main="1F4E79",red="C62828",amber="E65100",green="1B5E20",blue="0D47A1",
             gray="546E7A",orange="BF360C",teal="006064",gold="F57F17",
             red_l="FFEBEE",amber_l="FFF3E0",green_l="E8F5E9",blue_l="E3F2FD",
             gray_l="ECEFF1",orange_l="FBE9E7",gold_l="FFFDE7",low_l="FCE4EC",
             row1="EBF3FB",row2="F2F9EC",white="FFFFFF")
    ST = {'Критично':(C['red'],C['red_l']),'Низько':(C['amber'],C['amber_l']),
          'Норма':(C['green'],C['green_l']),'Надлишок':(C['blue'],C['blue_l'])}
    AB    = {'A':C['green'],'B':C['blue'],'?':C['gray']}
    AB_MI = {'A':C['gold'],'B':C['blue'],'C':C['gray'],'?':C['gray']}
    TR    = {'↑ зростає':C['green'],'↑ новий':C['teal'],'↓ спадає':C['red'],
             '→ стабільно':C['gray'],'—':C['gray']}

    wb = Workbook()
    p = (f"Маржа≥{params['mg_min']}% | Lead {params['lead']}д | Safety {params['safety']}д | "
         f"α={params['avail_alpha']} | λ={params['lambda_val']} | K={meta['season_K']:.3f}")

    def ws_init(ws, title, warn, cols, wc="7F3F00", wb_="FFF9C4"):
        lc = get_column_letter(len(cols))
        ws.merge_cells(f"A1:{lc}1"); ws["A1"] = title
        ws["A1"].font=hf(sz=9); ws["A1"].fill=fl(C['main'])
        ws["A1"].alignment=la(); ws.row_dimensions[1].height=20
        ws.merge_cells(f"A2:{lc}2"); ws["A2"] = warn
        ws["A2"].font=Font(name="Arial",size=9,italic=True,color=wc)
        ws["A2"].fill=fl(wb_); ws["A2"].alignment=la(); ws.row_dimensions[2].height=18
        for col,(h,w) in enumerate(cols,1):
            c=ws.cell(row=3,column=col,value=h)
            c.font=hf(sz=9); c.fill=fl(C['main']); c.alignment=ca(); c.border=tb()
            ws.column_dimensions[get_column_letter(col)].width=w
        ws.row_dimensions[3].height=40

    # Sheet 1 — загальна потреба до розподілу
    ws1 = wb.active; ws1.title = "Загальна потреба"
    COLS1 = [("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
             ("Маржа %",10),("MI/день",11),("Залишок+\nтранзит",10),
             ("Дата нуля",11),("Дні\nдо нуля",9),("Тренд",11),
             ("Замовити\n14+30д",12),("Замовити 60д\n(знижка)",12),("Сума замовл.\n(грн)",14)]
    ws_init(ws1, f"ЗАГАЛЬНА ПОТРЕБА | {today.strftime('%d.%m.%Y')} | {p}",
            f"⚑ = наявність <{params['low_avail']}%. "
            f"Золото ABC_MI=A (топ 70% маржинального доходу/день).", COLS1)

    order = sorted([r for r in results if r['rec']>0 or r['rec_60']>0],
                   key=lambda x: x['days_left'])
    for ri, r in enumerate(order):
        row = ri+4; dl = r['days_left']; low = r.get('low_avail', False)
        _, bg = ST.get(r['status'], (C['gray'],C['gray_l']))
        if low: bg = C['low_l']
        elif r['status']=='Норма': bg = C['row1'] if ri%2==0 else C['row2']
        tr = r.get('trend','—'); mi = r.get('mi_day'); abc_mi = r.get('abc_mi','?')
        sd_ = r.get('safety_disc',0)
        buy_p = r.get('buy_price') or 0
        qty_for_sum = r['rec_60'] if r['rec_60']>0 else r['rec']
        order_sum = round(qty_for_sum * buy_p, 2) if qty_for_sum > 0 and buy_p > 0 else None
        vals = [r['sku'],r['name'],r['abc'],abc_mi,r['avg_margin'],mi,r['eff_stock'],
                r.get('zero_date'),dl if dl<999 else None,tr,r['rec'],
                r['rec_60'] if r['rec_60']>0 else None,
                order_sum]
        for col, val in enumerate(vals, 1):
            c = ws1.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if col==1:
                c.value=("⚑ " if low else "")+str(r['sku'])
                c.font=Font(name="Arial",bold=True,size=10,color="880E4F" if low else "000000"); c.alignment=la()
            elif col==2: c.font=cf(sz=9); c.alignment=la()
            elif col==3: c.font=Font(name="Arial",bold=True,size=11,color=AB.get(r['abc'],C['gray'])); c.alignment=ca()
            elif col==4: c.font=Font(name="Arial",bold=True,size=11,color=AB_MI.get(abc_mi,C['gray'])); c.alignment=ca()
            elif col==5:
                if val: c.number_format='0.0"%"'
                c.font=Font(name="Arial",bold=True,size=10,color=C['green'] if (val or 0)>=params['mg_a'] else C['blue']); c.alignment=ca()
            elif col==6:
                if val: c.number_format='#,##0.00'
                ta_=meta.get('ta') or 0; tb__=meta.get('tb') or 0
                c.font=Font(name="Arial",bold=True,size=10,
                            color=C['gold'] if (val or 0)>=ta_ else C['blue'] if (val or 0)>=tb__ else C['gray']); c.alignment=ca()
            elif col==7: c.font=cf(bold=True,sz=10); c.alignment=ca()
            elif col==8:
                if val:
                    c.font=Font(name="Arial",bold=True,size=10,color=C['red'] if dl<5 else C['amber'] if dl<15 else "000000")
                else: c.value="∞"; c.font=cf(sz=10,color="AAAAAA")
                c.alignment=ca()
            elif col==9:
                c.font=Font(name="Arial",bold=True,size=11,
                            color=C['red'] if dl<5 else C['amber'] if dl<15 else C['green']); c.alignment=ca()
            elif col==10: c.font=Font(name="Arial",bold=True,size=10,color=TR.get(tr,C['gray'])); c.alignment=ca()
            elif col==11:
                if val and val>0: c.font=Font(name="Arial",bold=True,size=12,color=C['main'])
                else: c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
            elif col==12:
                if val and val>0:
                    c.font=Font(name="Arial",bold=True,size=12,color=C['amber'])
                    c.fill=fl(C['amber_l'])
                else:
                    c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
            elif col==13:
                if val and val>0:
                    c.number_format='#,##0.00'
                    c.font=Font(name="Arial",bold=True,size=10,color=C['main'])
                else:
                    c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
                c.alignment=ca()
    ws1.freeze_panes="A4"; ws1.auto_filter.ref=f"A3:{get_column_letter(len(COLS1))}{len(order)+3}"
    # Підсумковий рядок суми
    sr = len(order)+4
    ws1.cell(row=sr,column=1,value="ВСЬОГО:").font=Font(name="Arial",bold=True,size=10,color=C['main'])
    ws1.cell(row=sr,column=11,value=sum(r['rec'] for r in order)).font=Font(name="Arial",bold=True,size=12,color=C['main'])
    td=sum(r['rec_60'] for r in order if r['rec_60']>0)
    ws1.cell(row=sr,column=12,value=td if td else None).font=Font(name="Arial",bold=True,size=12,color=C['amber'])
    ts_v=sum((r['rec_60'] if r['rec_60']>0 else r['rec'])*(r.get('buy_price') or 0) for r in order)
    cs_v=ws1.cell(row=sr,column=13,value=round(ts_v,2) if ts_v>0 else None)
    cs_v.font=Font(name="Arial",bold=True,size=12,color=C['main']); cs_v.number_format='#,##0.00' 

    # Sheet 2 — Аналіз SKU
    ws2 = wb.create_sheet("Аналіз SKU")
    COLS2 = [("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
             ("Маржа %",10),("MI/день",11),("Наявн %",8),("Тренд",10),
             ("avg/день",10),("Залишок",8),("Транзит",8),
             ("Дата нуля",11),("Дні\nдо нуля",9),("Замовити\n14+30д",12),("Замовити\n(знижка)",12)]
    ws_init(ws2, f"Аналіз SKU — {len(results)} регулярних | {today.strftime('%d.%m.%Y')} | {p}",
            f"ABC_MI: A=золото (топ 70% MI), B=синій, C=сірий. ⚑ = наявність <{params['low_avail']}%.", COLS2)
    for ri, r in enumerate(sorted(results, key=lambda x: x['days_left'])):
        row=ri+4; dl=r['days_left']; low=r.get('low_avail',False)
        _,bg=ST.get(r['status'],(C['gray'],C['gray_l']))
        if low: bg=C['low_l']
        elif r['status']=='Норма': bg=C['row1'] if ri%2==0 else C['row2']
        tr=r.get('trend','—'); mi=r.get('mi_day'); abc_mi=r.get('abc_mi','?'); sd_=r.get('safety_disc',0)
        rec60_str=f"{r['rec_60']} шт ({sd_}д)" if r['rec_60']>0 else "—"
        vals=[r['sku'],r['name'],r['abc'],abc_mi,r['avg_margin'],mi,r['avail_pct'],
              tr,r['avg_day'],r['stock'],r['transit'],r.get('zero_date'),
              dl if dl<999 else None,r['rec'],rec60_str]
        for col,val in enumerate(vals,1):
            c=ws2.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if col==1:
                c.value=("⚑ " if low else "")+str(r['sku'])
                c.font=Font(name="Arial",bold=True,size=10,color="880E4F" if low else "000000"); c.alignment=la()
            elif col==2: c.font=cf(sz=9); c.alignment=la()
            elif col==3: c.font=Font(name="Arial",bold=True,size=11,color=AB.get(r['abc'],C['gray'])); c.alignment=ca()
            elif col==4: c.font=Font(name="Arial",bold=True,size=11,color=AB_MI.get(abc_mi,C['gray'])); c.alignment=ca()
            elif col==5:
                if val: c.number_format='0.0"%"'
                c.font=Font(name="Arial",bold=True,size=10,color=C['green'] if (val or 0)>=params['mg_a'] else C['blue']); c.alignment=ca()
            elif col==6:
                if val: c.number_format='#,##0.00'
                c.font=cf(sz=10,color=C['gold'] if (val or 0)>=(meta.get('ta') or 0) else C['blue']); c.alignment=ca()
            elif col==7:
                if val: c.number_format='0"%"'
                bc=C['red'] if (val or 0)<params['low_avail'] else C['amber'] if (val or 0)<50 else C['green']
                c.font=Font(name="Arial",bold=low,size=10,color=bc); c.alignment=ca()
            elif col==8: c.font=Font(name="Arial",bold=True,size=10,color=TR.get(tr,C['gray'])); c.alignment=ca()
            elif col==9: c.number_format='0.0000'; c.font=Font(name="Arial",bold=True,size=10,color=C['blue']); c.alignment=ca()
            elif col in(10,11): c.font=cf(sz=10); c.alignment=ca()
            elif col==12:
                if val:
                    fc2,_=ST.get(r['status'],(C['gray'],C['gray_l']))
                    c.font=Font(name="Arial",bold=True,size=10,color=fc2)
                else: c.value="∞"; c.font=cf(sz=10,color="AAAAAA")
                c.alignment=ca()
            elif col==13:
                fc2,_=ST.get(r['status'],(C['gray'],C['gray_l']))
                c.font=Font(name="Arial",bold=True,size=11,color=fc2); c.alignment=ca()
            elif col==14:
                if val and val>0: c.font=Font(name="Arial",bold=True,size=11,color=C['main'])
                else: c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
            else: c.font=cf(sz=10); c.alignment=ca()
    ws2.freeze_panes="A4"; ws2.auto_filter.ref=f"A3:{get_column_letter(len(COLS2))}{len(results)+3}"

    # Sheet 3 — Ручне рішення
    ws3 = wb.create_sheet("Ручне рішення")
    low_rows = sorted([r for r in results if r.get('low_avail')], key=lambda x: x['avail_pct'])
    COLS3 = [("SKU",13),("Назва",42),("ABC",7),("Маржа %",10),
             ("Наявн %",9),("Тренд",11),("Залишок",8),("Дата нуля",11)]
    ws_init(ws3, f"Ручне рішення — наявність <{params['low_avail']}% | {len(low_rows)} SKU",
            "Прогноз ненадійний. Вирішіть вручну.", COLS3, "880E4F", "FCE4EC")
    for ri,r in enumerate(low_rows):
        row=ri+4; bg="FCE4EC" if ri%2==0 else C['white']; dl=r['days_left']; tr=r.get('trend','—')
        for col,val in enumerate([r['sku'],r['name'],r['abc'],r['avg_margin'],
                                   r['avail_pct'],tr,r['stock'],r.get('zero_date')],1):
            c=ws3.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if col==1: c.font=cf(bold=True,sz=10); c.alignment=la()
            elif col==2: c.font=cf(sz=9); c.alignment=la()
            elif col==3: c.font=Font(name="Arial",bold=True,size=11,color=AB.get(r['abc'],C['gray'])); c.alignment=ca()
            elif col==4:
                if val: c.number_format='0.0"%"'
                c.font=Font(name="Arial",bold=True,size=10,color=C['blue']); c.alignment=ca()
            elif col==5:
                if val: c.number_format='0"%"'
                c.font=Font(name="Arial",bold=True,size=10,color=C['red']); c.alignment=ca()
            elif col==6: c.font=Font(name="Arial",bold=True,size=10,color=TR.get(tr,C['gray'])); c.alignment=ca()
            elif col==7: c.font=cf(sz=10); c.alignment=ca()
            elif col==8:
                if val:
                    c.font=Font(name="Arial",size=10,color=C['red'] if dl<5 else C['amber'] if dl<15 else "000000")
                else: c.value="∞"; c.font=cf(sz=10,color="AAAAAA")
                c.alignment=ca()

    # Sheet 4 — Разовий попит
    ws4 = wb.create_sheet("Разовий попит")
    COLS4 = [("SKU",13),("Назва",42),("Маржа %",10),("Чист. міс.",12),("Продано",12),("Причина",28)]
    ws_init(ws4, f"Разовий попит — {len(sporadic)} SKU",
            f"< {params['min_months']} міс. АБО < {params['min_qty']} шт — виключені з замовлень.",
            COLS4, "BF360C", "FBE9E7")
    for ri,r in enumerate(sorted(sporadic,key=lambda x:-(x.get('avg_margin') or 0))[:300]):
        row=ri+4; bg=C['orange_l'] if ri%2==0 else C['white']
        cmo=sum(1 for lbl in (data['meta']['months'] if 'meta' in data else []))
        ts=r.get('total_clean_sold', 0)
        for col,val in enumerate([r['sku'],r['name'],r.get('avg_margin'),
                                   r.get('clean_months',0),ts,r.get('sporadic_reason','')],1):
            c=ws4.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if col==1: c.font=cf(bold=True,sz=10); c.alignment=la()
            elif col==2: c.font=cf(sz=9); c.alignment=la()
            elif col==3:
                if val: c.number_format='0.0"%"'
                c.font=cf(sz=10,color=C['blue']); c.alignment=ca()
            elif col==6: c.font=Font(name="Arial",size=10,italic=True,color=C['orange']); c.alignment=la()
            else: c.font=cf(sz=10); c.alignment=ca()

    # Окремі замовлення постачальникам
    supplier_orders = data.get('supplier_orders', {})
    supplier_configs = data.get('supplier_configs', [])
    supplier_cols = [
        ("Артикул",15),("Код у постачальника",20),("Назва",42),("Кількість",11),("Ціна прайсу",13),
        ("Ціна з кориг.",14),("Сума",14),("Знижка %",10),("Маржа %",10),
        ("Lead time",10),("Дні до 0",10),("Причина вибору",44),
    ]
    used_sheet_names = set(wb.sheetnames)
    for supplier_index, supplier_config in enumerate(supplier_configs, 1):
        supplier = supplier_config['name']
        lines = supplier_orders.get(supplier, [])
        raw_name = re.sub(r'[\\/*?:\[\]]', '_', f"Замовл {supplier_index} {supplier}")[:31]
        sheet_name = raw_name or f"Замовлення {supplier_index}"
        suffix = 2
        while sheet_name in used_sheet_names:
            sheet_name = f"{raw_name[:28]} {suffix}"[:31]
            suffix += 1
        used_sheet_names.add(sheet_name)
        ws_supplier = wb.create_sheet(sheet_name)
        ws_init(
            ws_supplier,
            f"ЗАМОВЛЕННЯ — {supplier} | {today.strftime('%d.%m.%Y')}",
            f"SKU: {len(lines)} | Кількість: {sum(line['quantity'] for line in lines)} | "
            f"Відома сума: {sum(line['order_value'] for line in lines):,.2f} | "
            f"Без ціни: {sum(not line.get('price_known', True) for line in lines)} SKU",
            supplier_cols,
        )
        for row_index, line in enumerate(lines, 4):
            price_known = line.get('price_known', True)
            values = [
                line['sku'], line.get('supplier_sku', line['sku']), line['name'], line['quantity'], line['unit_price'],
                line['landed_unit_price'], line['order_value'] if price_known else None, line['discount_pct'],
                line['margin_pct'], line['lead_time_days'], line['days_left'],
                line.get('selection_reason', ''),
            ]
            for column_index, value in enumerate(values, 1):
                cell = ws_supplier.cell(row=row_index, column=column_index, value=value)
                cell.fill = fl(C['row1'] if row_index % 2 == 0 else C['white'])
                cell.border = tb(); cell.font = cf(bold=column_index in (1,4,7), sz=10)
                cell.alignment = la() if column_index in (1,2,3,12) else ca()
                if column_index in (5,6,7): cell.number_format = '#,##0.00'
                if column_index in (8,9) and value is not None: cell.number_format = '0.0"%"'
        total_row = len(lines) + 4
        ws_supplier.cell(total_row, 1, "ВСЬОГО").font = hf(C['main'], sz=10)
        ws_supplier.cell(total_row, 4, sum(line['quantity'] for line in lines)).font = hf(C['main'], sz=10)
        total_cell = ws_supplier.cell(total_row, 7, round(sum(line['order_value'] for line in lines), 2))
        total_cell.font = hf(C['main'], sz=10); total_cell.number_format = '#,##0.00'
        ws_supplier.freeze_panes = "A4"
        ws_supplier.auto_filter.ref = f"A3:L{max(3, len(lines)+3)}"

    # SKU, які не вдалося замовити в жодного постачальника
    unallocated = data.get('unallocated', [])
    ws_unallocated = wb.create_sheet("Нерозподілені SKU")
    unallocated_cols = [
        ("SKU",13),("Назва",42),("Потрібно",11),("Не розподілено",15),
        ("Дні до 0",10),("Причина",48),
    ]
    ws_init(
        ws_unallocated,
        f"Нерозподілені SKU — {len(unallocated)}",
        "Перевірте наявність, ціну, MOQ, ліміти, lead time та мінімальну маржу.",
        unallocated_cols,
        "880E4F", "FCE4EC",
    )
    for row_index, line in enumerate(unallocated, 4):
        values = [line['sku'], line['name'], line['requested_qty'], line['unallocated_qty'],
                  line['days_left'], line['reason']]
        for column_index, value in enumerate(values, 1):
            cell = ws_unallocated.cell(row=row_index, column=column_index, value=value)
            cell.fill = fl(C['low_l'] if row_index % 2 == 0 else C['white'])
            cell.border = tb(); cell.font = cf(bold=column_index in (1,4), sz=10)
            cell.alignment = la() if column_index in (1,2,6) else ca()
    ws_unallocated.freeze_panes = "A4"
    ws_unallocated.auto_filter.ref = f"A3:F{max(3, len(unallocated)+3)}"

    # Аудит кодів другого постачальника, які не вдалося зіставити
    mapping_audits = data.get('barcode_mapping_audits', {})
    audit_rows = []
    for supplier, audit in mapping_audits.items():
        for code in audit.get('unknown_codes', []):
            audit_rows.append([supplier, "Не знайдено відповідність", code, ""])
        for duplicate in audit.get('duplicate_articles', []):
            audit_rows.append([
                supplier, "Кілька кодів одного артикулу",
                ", ".join(duplicate.get('supplier_codes', [])), duplicate.get('article', ""),
            ])
    ws_codes = wb.create_sheet("Контроль штрихкодів")
    code_cols = [("Постачальник",22),("Статус",31),("Код у прайсі",25),("Артикул",18)]
    ws_init(
        ws_codes,
        f"КОНТРОЛЬ ШТРИХКОДІВ — {len(audit_rows)} записів для перевірки",
        "Невідомі коди не включаються до замовлення, доки для них не додано відповідність.",
        code_cols,
        "880E4F", "FCE4EC",
    )
    for row_index, values in enumerate(audit_rows, 4):
        for column_index, value in enumerate(values, 1):
            cell = ws_codes.cell(row=row_index, column=column_index, value=value)
            cell.fill = fl(C['low_l'] if row_index % 2 == 0 else C['white'])
            cell.border = tb(); cell.font = cf(bold=column_index in (2,3), sz=10)
            cell.alignment = la()
    ws_codes.freeze_panes = "A4"
    ws_codes.auto_filter.ref = f"A3:D{max(3, len(audit_rows)+3)}"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── UI ───────────────────────────────────────
st.title("📦 ProcureAI — Аналіз закупівель")
st.caption("Розрахунок потреби → доступне замовляємо в ДСН → решту замовляємо в iHerb")

with st.sidebar:
    st.header("⚙️ Параметри")
    with st.expander("💰 Маржа та ABC", expanded=True):
        mg_min = st.slider("Мінімальна маржа (%)", 5, 50, 35)
        mg_a   = st.slider("A-клас: маржа ≥ (%)", mg_min, 70, max(mg_min+10, 50))
        a_mult = st.slider("Бонус запасу A-класу (×)", 1.0, 2.0, 1.3, 0.05)
    with st.expander("📅 Замовлення", expanded=True):
        lead    = st.slider("Lead time (дні)", 1, 30, 14)
        safety  = st.slider("Страховий запас (дні)", 7, 60, 30)
        st.markdown("**Знижки постачальника:**")
        disc_thr  = st.slider("Поріг знижки 1 (%)", 5, 60, 10, help="Знижка ≥ цього % → запас на N днів")
        safety60  = st.slider("Запас при знижці 1 (дні)", 30, 90, 60)
        disc_thr2 = st.slider("Поріг знижки 2 (%)", 15, 70, 40, help="Знижка ≥ цього % → більший запас")
        safety90  = st.slider("Запас при знижці 2 (дні)", 60, 180, 90)
        if disc_thr2 <= disc_thr:   disc_thr2 = disc_thr + 1
        if safety90  <= safety60:   safety90  = safety60 + 1
    with st.expander("📊 Попит", expanded=True):
        min_months = st.slider("Мін. місяців з продажами", 2, 9, 6)
        min_qty    = st.slider("Мін. продано штук", 1, 30, 12)
        low_avail  = st.slider("Поріг низької наявності (%)", 10, 40, 20)
    with st.expander("🔢 Коефіцієнти", expanded=False):
        avail_alpha = st.slider("Коефіцієнт наявності α", 0.3, 1.5, 0.7, 0.05,
                                help="avail_K = (наявн%)^α")
        lambda_val  = st.slider("Часові ваги λ", 0.05, 0.5, 0.25, 0.05,
                                help="Більше λ = більша вага останніх місяців")

    with st.expander("⚖️ Правила вибору постачальника", expanded=True):
        strategy_label = st.selectbox(
            "Стратегія",
            ["Спочатку ДСН, решта — iHerb", "Пріоритет у межах допуску", "Найнижча ціна", "Баланс ціни та строку"],
            help="Основне правило: усе доступне беремо в ДСН, решту включаємо в замовлення iHerb.",
        )
        strategy = {
            "Спочатку ДСН, решта — iHerb": "secondary_first_primary_fallback",
            "Пріоритет у межах допуску": "priority_within_tolerance",
            "Найнижча ціна": "lowest_price",
            "Баланс ціни та строку": "balanced",
        }[strategy_label]
        if strategy == "secondary_first_primary_fallback":
            st.info("ДСН: беремо доступні позиції. iHerb: автоматично отримує весь залишок, навіть без ціни в прайсі.")
            price_tolerance = 0
        else:
            price_tolerance = st.slider("Допустима різниця ціни (%)", 0, 30, 5)
        min_purchase_margin = st.slider("Мін. маржа після вибору прайсу (%)", 0, 60, 0)
        max_supplier_lead = st.slider("Макс. lead time постачальника (дні, 0 = без межі)", 0, 120, 0)

    with st.expander("🏪 Постачальник 1", expanded=False):
        supplier_1_name = st.text_input("Назва", "iHerb", key="supplier_1_name")
        supplier_1_priority = st.number_input("Пріоритет", 1, 10, 1, key="supplier_1_priority")
        supplier_1_lead = st.number_input("Lead time (дні)", 0, 180, int(lead), key="supplier_1_lead")
        supplier_1_adjustment = st.number_input(
            "Коригування ціни: доставка/мито (%)", -50.0, 200.0, 0.0, 0.5, key="supplier_1_adjustment")
        supplier_1_moq = st.number_input("MOQ за SKU (шт)", 1, 10000, 1, key="supplier_1_moq")
        supplier_1_min_value = st.number_input(
            "Мінімальна сума замовлення (0 = немає)", 0.0, 100000000.0, 0.0, 100.0, key="supplier_1_min_value")
        supplier_1_limit = st.number_input(
            "Ліміт замовлення (0 = без ліміту)", 0.0, 100000000.0, 0.0, 100.0, key="supplier_1_limit")

    with st.expander("🏬 Постачальник 2", expanded=False):
        supplier_2_name = st.text_input("Назва", "ДСН", key="supplier_2_name")
        supplier_2_source = st.selectbox(
            "Джерело прайсу", ["Google Sheets (автоматично)", "Excel-файл"],
            key="supplier_2_source")
        supplier_2_google_url = st.text_input(
            "Посилання Google Sheets",
            "https://docs.google.com/spreadsheets/d/1k0hGCK2LbgnzaYkqCuj7QXXRdDbdbvjbGhK9OjZGvG8/edit?gid=0#gid=0",
            key="supplier_2_google_url",
            help="Таблиця повинна мати доступ: усі, хто має посилання — читач.")
        supplier_2_priority = st.number_input("Пріоритет", 1, 10, 2, key="supplier_2_priority")
        supplier_2_lead = st.number_input("Lead time (дні)", 0, 180, 7, key="supplier_2_lead")
        supplier_2_adjustment = st.number_input(
            "Коригування ціни: доставка/мито (%)", -50.0, 200.0, 0.0, 0.5, key="supplier_2_adjustment")
        supplier_2_moq = st.number_input("MOQ за SKU (шт)", 1, 10000, 1, key="supplier_2_moq")
        supplier_2_min_value = st.number_input(
            "Мінімальна сума замовлення (0 = немає)", 0.0, 100000000.0, 0.0, 100.0, key="supplier_2_min_value")
        supplier_2_limit = st.number_input(
            "Ліміт замовлення (0 = без ліміту)", 0.0, 100000000.0, 0.0, 100.0, key="supplier_2_limit")

params = dict(mg_min=mg_min, mg_a=mg_a, a_mult=a_mult,
              lead=lead, safety=safety, safety60=safety60, safety90=safety90,
              disc_thr=disc_thr, disc_thr2=disc_thr2,
              min_months=min_months, min_qty=min_qty, low_avail=low_avail,
              avail_alpha=avail_alpha, lambda_val=lambda_val)

if supplier_2_name.strip() == supplier_1_name.strip():
    supplier_2_name = f"{supplier_2_name.strip()} 2"

supplier_configs = [
    dict(name=supplier_1_name.strip() or "iHerb", priority=supplier_1_priority,
         lead_time_days=supplier_1_lead, price_adjustment_pct=supplier_1_adjustment,
         min_order_qty=supplier_1_moq, min_order_value=supplier_1_min_value,
         order_limit_value=supplier_1_limit),
    dict(name=supplier_2_name.strip() or "ДСН", priority=supplier_2_priority,
         lead_time_days=supplier_2_lead, price_adjustment_pct=supplier_2_adjustment,
         min_order_qty=supplier_2_moq, min_order_value=supplier_2_min_value,
         order_limit_value=supplier_2_limit),
]
allocation_rules = dict(
    strategy=strategy,
    price_tolerance_pct=price_tolerance,
    min_purchase_margin_pct=min_purchase_margin,
    max_lead_time_days=max_supplier_lead,
    price_weight=0.7,
    lead_time_weight=0.2,
)

# ── Завантаження файлів ──
col1, col2, col3 = st.columns(3)
with col1:
    f_template = st.file_uploader(
        "📋 **Файл 1 — Шаблон**",
        type=["xlsx","xls"], key="template",
        help="Вкладки: 'продажі дані', 'наявність на складі', 'Залишки'")
with col2:
    f_prices_1 = st.file_uploader(
        f"💰 **Прайс — {supplier_configs[0]['name']}**",
        type=["xlsx","xls"], key="prices_1",
        help="Обов'язкові колонки: SKU/Артикул і Ціна. Порядок колонок довільний.")
with col3:
    if supplier_2_source == "Excel-файл":
        f_prices_2 = st.file_uploader(
            f"💰 **Прайс — {supplier_configs[1]['name']}**",
            type=["xlsx","xls"], key="prices_2",
            help="Підтримуються Наявність, Доступна кількість, Знижка, Lead time та MOQ.")
    else:
        f_prices_2 = None
        st.info(f"🔄 {supplier_configs[1]['name']}: прайс читається з Google Sheets кожні 5 хвилин")
        if st.button("Оновити онлайн-прайс зараз", use_container_width=True):
            load_google_sheet_prices.clear()

if f_template:
    with st.spinner("Читаємо шаблон..."):
        wb_t = load_wb(f_template.read())
        sku_data, months_labels, stock_map, avail_map, errors = parse_template(wb_t)

    if errors:
        for e in errors: st.error(e)
        st.stop()

    price_maps = {supplier_configs[0]['name']: {}, supplier_configs[1]['name']: {}}
    barcode_map, barcode_map_info = load_barcode_mapping()
    mapping_audits = {}
    if barcode_map_info.get("error"):
        st.warning(f"Таблиця штрихкодів: {barcode_map_info['error']}")
    else:
        st.caption(
            f"🔗 Таблиця відповідності: {barcode_map_info['unique_barcodes']:,} штрихкодів → "
            f"{barcode_map_info['unique_articles']:,} артикулів"
        )

    def show_price_result(supplier_name, offers, price_warnings, source_label, map_barcodes=False):
        if map_barcodes and barcode_map:
            offers, mapping_audit = apply_barcode_mapping(offers, barcode_map)
            mapping_audits[supplier_name] = mapping_audit
            if mapping_audit['mapped_by_barcode']:
                price_warnings.append(
                    f"{supplier_name}: зіставлено {mapping_audit['mapped_by_barcode']} штрихкодів з артикулами"
                )
            if mapping_audit['unknown_codes']:
                examples = ", ".join(mapping_audit['unknown_codes'][:8])
                price_warnings.append(
                    f"{supplier_name}: не знайдено {len(mapping_audit['unknown_codes'])} кодів "
                    f"у таблиці відповідності ({examples})"
                )
        price_maps[supplier_name] = offers
        st.success(
            f"✅ {supplier_name} ({source_label}): {len(offers)} SKU, "
            f"в наявності {sum(1 for offer in offers.values() if offer['in_stock'])}")
        for warning in price_warnings:
            st.warning(warning)

    if f_prices_1:
        with st.spinner(f"Читаємо прайс: {supplier_configs[0]['name']}..."):
            wb_p = load_wb(f_prices_1.read())
            offers, price_warnings = parse_prices(wb_p, supplier_configs[0]['name'])
        show_price_result(supplier_configs[0]['name'], offers, price_warnings, "Excel")

    if supplier_2_source == "Google Sheets (автоматично)":
        try:
            with st.spinner(f"Оновлюємо онлайн-прайс: {supplier_configs[1]['name']}..."):
                offers, price_warnings = load_google_sheet_prices(
                    supplier_2_google_url, supplier_configs[1]['name'])
            show_price_result(
                supplier_configs[1]['name'], offers, price_warnings, "Google Sheets", map_barcodes=True)
        except ValueError as exc:
            st.error(f"❌ Не вдалося прочитати онлайн-прайс: {exc}")
            st.caption("У Google Sheets встановіть доступ: Усі, хто має посилання → Читач.")
    elif f_prices_2:
        with st.spinner(f"Читаємо прайс: {supplier_configs[1]['name']}..."):
            wb_p = load_wb(f_prices_2.read())
            offers, price_warnings = parse_prices(wb_p, supplier_configs[1]['name'])
        show_price_result(supplier_configs[1]['name'], offers, price_warnings, "Excel", map_barcodes=True)
    if not any(price_maps.values()):
        st.info("💡 Прайси не завантажено — потреба буде розрахована, але SKU залишаться нерозподіленими")

    # Для прогнозу ціни/знижки потрібна одна референсна пропозиція; фактичний
    # постачальник визначається окремим алгоритмом нижче.
    price_map = build_reference_price_map(price_maps)

    # Зведення по вкладках
    with st.expander("🔍 Статистика завантажених даних", expanded=False):
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("SKU в продажах", len(sku_data))
        c2.metric("SKU в наявності", len(stock_map))
        c3.metric("SKU в залишках", len(avail_map))
        c4.metric("Місяців даних", len(months_labels) if months_labels else 0)
        st.write(f"**Місяці:** {', '.join(months_labels) if months_labels else '—'}")

    with st.spinner("Аналізуємо..."):
        data = run_analysis(sku_data, months_labels, stock_map, avail_map, price_map, params)

    demand_lines = [
        dict(
            sku=row['sku'], name=row['name'], days_left=row['days_left'],
            order_qty=row['rec_60'] if row.get('rec_60', 0) > 0 else row['rec'],
            sell_price=row.get('sell_price'),
        )
        for row in data['regular']
        if row['rec'] > 0 or row.get('rec_60', 0) > 0
    ]
    allocation = allocate_orders(demand_lines, price_maps, supplier_configs, allocation_rules)
    data['supplier_orders'] = allocation['orders']
    data['supplier_summary'] = allocation['summary']
    data['unallocated'] = allocation['unallocated']
    data['excluded_suppliers'] = allocation['excluded_suppliers']
    data['supplier_configs'] = supplier_configs
    data['allocation_rules'] = allocation_rules
    data['barcode_mapping_audits'] = mapping_audits

    meta     = data['meta']
    regular  = data['regular']
    sporadic = data['sporadic']
    excl_mg  = data['excl_mg']
    total    = len(regular)+len(sporadic)+len(excl_mg)

    # KPI
    st.subheader("📊 Зведення")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Всього SKU", total)
    c2.metric("Постійний попит", len(regular))
    c3.metric("Разовий попит", len(sporadic))
    c4.metric("Критично (<5дн)", sum(1 for r in regular if r['status']=='Критично'))
    c5.metric("Розподілено, шт", sum(v['quantity'] for v in data['supplier_summary'].values()))
    c6.metric("Нерозподілено SKU", len(data['unallocated']))

    st.subheader("📅 Дата нуля залишків")
    buckets = [("🔴 Вже 0",0,1),("🔴 1-7 дн",1,8),("🟡 8-14 дн",8,15),
               ("🟡 15-30 дн",15,31),("🟢 31-60 дн",31,61),("🔵 >60 дн",61,999)]
    cols_b = st.columns(len(buckets))
    for i,(lbl,lo,hi) in enumerate(buckets):
        cols_b[i].metric(lbl, sum(1 for r in regular if lo<=r['days_left']<hi))

    tab1,tab2,tab3,tab4,tab5 = st.tabs([
        "🏪 Розподіл", "🛒 Загальна потреба", "📋 Аналіз SKU",
        "⚠️ Ручне рішення", "📦 Разовий попит",
    ])

    with tab1:
        if data['supplier_summary']:
            summary_cols = st.columns(len(data['supplier_summary']))
            for index, (supplier, summary) in enumerate(data['supplier_summary'].items()):
                unknown_note = f" · без ціни {summary.get('unknown_price_sku', 0)}" if summary.get('unknown_price_sku') else ""
                summary_cols[index].metric(
                    supplier,
                    f"{summary['quantity']} шт",
                    f"{summary['sku_count']} SKU · відома сума {summary['order_value']:,.2f}{unknown_note}",
                )
        for supplier in [item['name'] for item in supplier_configs]:
            lines = data['supplier_orders'].get(supplier, [])
            st.markdown(f"#### {supplier}")
            if not lines:
                st.info("Немає позицій для цього постачальника")
                continue
            supplier_df = pd.DataFrame([{
                'Артикул': line['sku'],
                'Код у постачальника': line.get('supplier_sku', line['sku']),
                'Назва': line['name'][:60],
                'Кількість': line['quantity'],
                'Ціна прайсу': line['unit_price'],
                'Ціна з коригуванням': line['landed_unit_price'],
                'Сума': line['order_value'] if line.get('price_known', True) else None,
                'Маржа %': line['margin_pct'],
                'Lead time': line['lead_time_days'],
                'Дні до 0': line['days_left'],
                'Причина вибору': line.get('selection_reason', ''),
            } for line in lines])
            st.dataframe(supplier_df, use_container_width=True, hide_index=True)

        if data['excluded_suppliers']:
            for excluded in data['excluded_suppliers']:
                st.warning(
                    f"{excluded['supplier']} виключено: сума {excluded['order_value']:,.2f} "
                    f"менша за мінімальне замовлення {excluded['min_order_value']:,.2f}")

        if data['unallocated']:
            st.markdown("#### ⚠️ Нерозподілені SKU")
            unallocated_df = pd.DataFrame([{
                'SKU': line['sku'], 'Назва': line['name'][:60],
                'Потрібно': line['requested_qty'], 'Не розподілено': line['unallocated_qty'],
                'Дні до 0': line['days_left'], 'Причина': line['reason'],
            } for line in data['unallocated']])
            st.dataframe(unallocated_df, use_container_width=True, hide_index=True)
        elif demand_lines:
            st.success("Усю потребу розподілено між постачальниками")

    with tab2:
        order = sorted([r for r in regular if r['rec']>0 or r['rec_60']>0],
                       key=lambda x: x['days_left'])
        if order:
            df = pd.DataFrame([{
                'SKU':       ('⚑ ' if r['low_avail'] else '')+r['sku'],
                'Назва':     r['name'][:50],
                'ABC':       r['abc'], 'ABC_MI': r.get('abc_mi','?'),
                'Маржа %':   r['avg_margin'],
                'MI/день':   r.get('mi_day'),
                'Залишок':   r['eff_stock'],
                'Дата нуля': r.get('zero_date','∞'),
                'Дні до 0':  r['days_left'] if r['days_left']<999 else '∞',
                'Тренд':     r['trend'],
                'Замовити 14+30': r['rec'],
                'Замовити 60д (знижка)': r['rec_60'] if r.get('rec_60',0)>0 else None,
                'Сума замовл. (грн)': round((r['rec_60'] if r.get('rec_60',0)>0 else r['rec'])*(r.get('buy_price') or 0),2) if (r.get('buy_price') or 0)>0 else None,
                'Знижка %':  r.get('price_disc',0) or '—',
            } for r in order])
            st.dataframe(df, use_container_width=True, height=500)
            disc1 = [r for r in order if r.get('rec_60',0)>0 and not r.get('use_90')]
            disc2 = [r for r in order if r.get('rec_60',0)>0 and r.get('use_90')]
            st.caption(
                f"Всього: {len(order)} SKU | стандарт: {sum(r['rec'] for r in order)} шт | "
                f"знижка 1 ({params['safety60']}д): {sum(r.get('rec_60',0) for r in disc1)} шт / {len(disc1)} SKU | "
                f"знижка 2 ({params.get('safety90',90)}д): {sum(r.get('rec_60',0) for r in disc2)} шт / {len(disc2)} SKU")
        else:
            st.success("Всі залишки в нормі — замовлення не потрібні")

    with tab3:
        df2 = pd.DataFrame([{
            'SKU':       ('⚑ ' if r['low_avail'] else '')+r['sku'],
            'Назва':     r['name'][:50],
            'ABC':       r['abc'], 'ABC_MI': r.get('abc_mi','?'),
            'Маржа %':   r['avg_margin'],
            'MI/день':   r.get('mi_day'),
            'Наявн %':   r['avail_pct'],
            'Тренд':     r['trend'],
            'avg/день':  r['avg_day'],
            'Залишок':   r['stock'],
            'Транзит':   r['transit'],
            'Дата нуля': r.get('zero_date','∞'),
            'Дні до 0':  r['days_left'] if r['days_left']<999 else '∞',
            'Статус':    r['status'],
            'Замовити':  r['rec'],
        } for r in sorted(regular, key=lambda x: x['days_left'])])
        st.dataframe(df2, use_container_width=True, height=500)

    with tab4:
        low_rows = sorted([r for r in regular if r.get('low_avail')], key=lambda x: x['avail_pct'])
        if low_rows:
            df3 = pd.DataFrame([{'SKU':r['sku'],'Назва':r['name'][:50],'ABC':r['abc'],
                                  'Маржа %':r['avg_margin'],'Наявн %':r['avail_pct'],
                                  'Тренд':r['trend'],'Дата нуля':r.get('zero_date','∞')}
                                 for r in low_rows])
            st.warning(f"⚑ {len(low_rows)} SKU мали наявність <{params['low_avail']}%")
            st.dataframe(df3, use_container_width=True, height=400)
        else:
            st.success("Немає SKU з низькою наявністю")

    with tab5:
        if sporadic:
            df4 = pd.DataFrame([{'SKU':r['sku'],'Назва':r['name'][:50],
                                  'Маржа %':r.get('avg_margin'),'Причина':r.get('sporadic_reason','')}
                                 for r in sorted(sporadic, key=lambda x:-(x.get('avg_margin') or 0))[:300]])
            st.info(f"🔵 {len(sporadic)} SKU з нерегулярним попитом")
            st.dataframe(df4, use_container_width=True, height=400)
        else:
            st.success("Всі SKU мають постійний попит")

    st.divider()
    with st.spinner("Генеруємо Excel..."):
        excel_buf = gen_excel(data, params)
    st.download_button(
        "📥 Завантажити Excel-звіт з двома замовленнями",
        data=excel_buf,
        file_name=f"ProcureAI_{date.today().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

else:
    st.info("👆 Завантажте Файл 1 (Шаблон) щоб почати аналіз")
    with st.expander("📖 Структура файлів"):
        st.markdown("""
**Файл 1 — Шаблон** (оновлюєте щодня):
| Вкладка | Структура |
|---|---|
| `продажі дані` | Блоки по 5 колонок на місяць: Назва | Артикул | Кількість | Рентаб% | _ |
| `наявність на складі` | Номенклатура | Артикул | Залишок | Замовлено у постачальників |
| `Залишки` | Блоки по 4 колонки на місяць: Артикул | Назва | Дні на складі | _ |

**Прайси постачальників 1 і 2** (окремі файли Excel):
- Обов'язково: `SKU`/`Артикул` та `Ціна`/`Цена`/`Price`
- Необов'язково: `Назва`, `Наявність`, `Доступна кількість`, `Знижка`, `Lead time`, `MOQ`
- Порядок колонок довільний; старий формат каталогу iHerb також підтримується.
        """)
