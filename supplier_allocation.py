"""Parsing supplier price lists and allocating purchase demand.

The module intentionally has no Streamlit dependency so the business rules can be
tested and changed without touching the UI or the sales forecast algorithm.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Any, Iterable


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def safe_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if math.isfinite(float(value)) else None
    text = clean_text(value).replace(" ", "").replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def parse_percent(value: Any) -> float:
    number = safe_number(value)
    if number is None:
        return 0.0
    if 0 < abs(number) < 1:
        number *= 100
    return round(max(number, 0), 2)


TRUE_WORDS = (
    "true", "yes", "y", "1", "так", "да", "є", "наяв", "в наявності",
    "в наличии", "available", "in stock",
)
FALSE_WORDS = (
    "false", "no", "n", "0", "ні", "нет", "немає", "відсут", "out of stock",
)


def parse_availability(value: Any, quantity: float | None = None) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) > 0
    text = clean_text(value).lower()
    if any((text == word if len(word) <= 3 else word in text) for word in FALSE_WORDS):
        return False
    if any((text == word if len(word) <= 3 else word in text) for word in TRUE_WORDS):
        return True
    return quantity is None or quantity > 0


HEADER_ALIASES = {
    "sku": ("sku", "артикул", "код товару", "код товара", "item code", "product code"),
    "barcode": ("штрихкод", "штрих-код", "штрих код", "barcode", "ean", "upc"),
    "name": ("назва", "название", "найменування", "наименование", "name", "product"),
    "price": (
        "ціна", "цена", "price", "cost", "закупівельна ціна", "закупочная цена",
        "purchase price", "unit price",
    ),
    "availability": ("наявність", "наличие", "availability", "in stock", "статус", "status"),
    "available_qty": (
        "доступна кількість", "доступное количество", "доступно", "залишок", "остаток",
        "stock qty", "available qty", "quantity", "qty", "кількість", "количество",
    ),
    "discount_pct": ("знижка", "скидка", "discount", "discount pct", "discount %"),
    "lead_time_days": (
        "lead time", "lead_time", "термін поставки", "срок поставки", "дні доставки",
        "дней доставки", "delivery days",
    ),
    "min_order_qty": (
        "moq", "мінімальна кількість", "минимальное количество", "мін замовлення",
        "мин заказ", "minimum order qty", "min qty",
    ),
}

BARCODE_HEADER_ALIASES = ("штрихкод", "штрих-код", "штрих код", "barcode", "ean", "upc")
ARTICLE_HEADER_ALIASES = ("артикул", "sku", "код товару", "код товара", "article")


def _normalise_header(value: Any) -> str:
    text = clean_text(value).lower().replace("_", " ")
    return re.sub(r"\s+", " ", text)


def _field_for_header(value: Any) -> str | None:
    header = _normalise_header(value)
    if not header:
        return None
    for field, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            alias_norm = _normalise_header(alias)
            if header == alias_norm or (len(alias_norm) >= 5 and alias_norm in header):
                return field
    return None


def normalise_product_code(value: Any) -> str:
    """Return a stable text key without dropping leading zeroes from barcodes."""

    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        return str(int(value)) if value.is_integer() else format(value, "f").rstrip("0").rstrip(".")

    text = clean_text(value).strip("'\"")
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def parse_barcode_mapping_rows(rows: Iterable[Iterable[Any]]) -> tuple[dict[str, str], dict]:
    """Parse a barcode-to-article table and report skipped or ambiguous rows."""

    values = [tuple(row) for row in rows]
    header_index = None
    article_column = None
    barcode_column = None
    for row_index, row in enumerate(values[:10]):
        for column_index, value in enumerate(row):
            header = _normalise_header(value)
            if header in ARTICLE_HEADER_ALIASES and article_column is None:
                article_column = column_index
            if header in BARCODE_HEADER_ALIASES and barcode_column is None:
                barcode_column = column_index
        if article_column is not None and barcode_column is not None:
            header_index = row_index
            break

    if header_index is None:
        return {}, {
            "total_rows": 0, "valid_pairs": 0, "empty_rows": 0, "conflicts": [],
            "error": "Не знайдено колонки «Артикул» і «Штрихкод»",
        }

    mapping: dict[str, str] = {}
    conflicts: list[dict] = []
    empty_rows = 0
    valid_pairs = 0
    article_barcodes: dict[str, set[str]] = {}
    for row in values[header_index + 1 :]:
        article = clean_text(row[article_column] if article_column < len(row) else None)
        barcode = normalise_product_code(row[barcode_column] if barcode_column < len(row) else None)
        if not article or not barcode:
            empty_rows += 1
            continue
        existing = mapping.get(barcode)
        if existing and existing != article:
            conflicts.append({"barcode": barcode, "articles": sorted({existing, article})})
            continue
        mapping[barcode] = article
        article_barcodes.setdefault(article, set()).add(barcode)
        valid_pairs += 1

    return mapping, {
        "total_rows": max(0, len(values) - header_index - 1),
        "valid_pairs": valid_pairs,
        "unique_barcodes": len(mapping),
        "unique_articles": len(article_barcodes),
        "articles_with_multiple_barcodes": sum(1 for codes in article_barcodes.values() if len(codes) > 1),
        "empty_rows": empty_rows,
        "conflicts": conflicts,
        "error": "",
    }


def apply_barcode_mapping(
    offers: dict[str, dict], barcode_map: dict[str, str]
) -> tuple[dict[str, dict], dict]:
    """Replace supplier barcodes with internal articles and keep an audit trail."""

    remapped: dict[str, dict] = {}
    mapped_count = 0
    direct_sku_count = 0
    unknown_codes: list[str] = []
    duplicate_articles: list[dict] = []

    for source_code, source_offer in offers.items():
        normalised_code = normalise_product_code(source_code)
        article = barcode_map.get(normalised_code)
        if article:
            mapped_count += 1
        elif re.search(r"[A-Za-zА-Яа-яІіЇїЄє]", normalised_code):
            article = clean_text(source_code)
            direct_sku_count += 1
        else:
            unknown_codes.append(normalised_code or clean_text(source_code))
            continue

        offer = dict(source_offer)
        offer["sku"] = article
        offer["supplier_sku"] = clean_text(source_code)
        offer["barcode"] = normalised_code if normalised_code in barcode_map else ""

        existing = remapped.get(article)
        if existing is not None:
            duplicate_articles.append({
                "article": article,
                "supplier_codes": [existing.get("supplier_sku", ""), offer["supplier_sku"]],
            })
            offer = min(
                [existing, offer],
                key=lambda item: (
                    0 if item.get("in_stock") else 1,
                    item.get("price", math.inf),
                    -(item.get("available_qty") or 0),
                ),
            )
        remapped[article] = offer

    return remapped, {
        "source_offers": len(offers),
        "mapped_by_barcode": mapped_count,
        "direct_sku": direct_sku_count,
        "unknown_codes": sorted(set(code for code in unknown_codes if code)),
        "duplicate_articles": duplicate_articles,
    }


def google_sheet_csv_url(sheet_url: str) -> str:
    """Convert a normal Google Sheets link into its public CSV export URL.

    Only the spreadsheet ID and numeric gid are retained, so arbitrary hosts or
    paths supplied through the UI are never requested by the application.
    """

    url = clean_text(sheet_url)
    match = re.search(r"docs\.google\.com/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    if not match:
        raise ValueError("Некоректне посилання Google Sheets")
    gid_match = re.search(r"(?:[?#&]|%23)gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv&gid={gid}"


def parse_supplier_prices(workbook: Any, supplier_name: str) -> tuple[dict[str, dict], list[str]]:
    """Parse the first worksheet using multilingual, order-independent headers.

    Missing optional columns are filled from supplier-level settings during allocation.
    For backward compatibility, the old iHerb layout falls back to fixed columns.
    """

    sheet = workbook[workbook.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return {}, [f"{supplier_name}: прайс порожній"]

    header_index = None
    columns: dict[str, int] = {}
    for row_index, row in enumerate(rows[:10]):
        candidate: dict[str, int] = {}
        for column_index, value in enumerate(row):
            field = _field_for_header(value)
            if field and field not in candidate:
                candidate[field] = column_index
        if ("sku" in candidate or "barcode" in candidate) and "price" in candidate:
            header_index = row_index
            columns = candidate
            break

    warnings: list[str] = []
    if header_index is None:
        # Legacy iHerb format: SKU | Name | Price | Old price | Availability | Discount
        header_index = 0
        columns = {"sku": 0, "name": 1, "price": 2, "availability": 4, "discount_pct": 5}
        warnings.append(f"{supplier_name}: заголовки не розпізнано, застосовано формат iHerb")

    def get(row: Iterable[Any], field: str) -> Any:
        values = tuple(row)
        index = columns.get(field)
        return values[index] if index is not None and index < len(values) else None

    offers: dict[str, dict] = {}
    invalid_price_count = 0
    for row in rows[header_index + 1 :]:
        # When both are present, prefer barcode: supplier articles may not match
        # the internal SKU, while barcode mapping is unambiguous.
        sku = clean_text(get(row, "barcode")) or clean_text(get(row, "sku"))
        if not sku or _normalise_header(sku) in (*HEADER_ALIASES["sku"], *HEADER_ALIASES["barcode"]):
            continue
        price = safe_number(get(row, "price"))
        if price is None or price <= 0:
            invalid_price_count += 1
            continue
        available_qty = safe_number(get(row, "available_qty"))
        if available_qty is not None:
            available_qty = max(0.0, available_qty)
        in_stock = parse_availability(get(row, "availability"), available_qty)
        offers[sku] = {
            "supplier": supplier_name,
            "sku": sku,
            "name": clean_text(get(row, "name")),
            "price": round(price, 4),
            "in_stock": in_stock,
            "available_qty": available_qty,
            "discount_pct": parse_percent(get(row, "discount_pct")),
            "lead_time_days": safe_number(get(row, "lead_time_days")),
            "min_order_qty": safe_number(get(row, "min_order_qty")),
        }

    if invalid_price_count:
        warnings.append(f"{supplier_name}: пропущено {invalid_price_count} рядків без коректної ціни")
    return offers, warnings


@dataclass(frozen=True)
class SupplierConfig:
    name: str
    enabled: bool = True
    priority: int = 1
    lead_time_days: float = 14
    price_adjustment_pct: float = 0
    min_order_qty: float = 1
    min_order_value: float = 0
    order_limit_value: float = 0

    @classmethod
    def from_mapping(cls, value: dict) -> "SupplierConfig":
        return cls(
            name=clean_text(value.get("name")) or "Постачальник",
            enabled=bool(value.get("enabled", True)),
            priority=max(1, int(value.get("priority", 1))),
            lead_time_days=max(0.0, float(value.get("lead_time_days", 14))),
            price_adjustment_pct=float(value.get("price_adjustment_pct", 0)),
            min_order_qty=max(0.0, float(value.get("min_order_qty", 1))),
            min_order_value=max(0.0, float(value.get("min_order_value", 0))),
            order_limit_value=max(0.0, float(value.get("order_limit_value", 0))),
        )


def build_reference_price_map(price_maps: dict[str, dict[str, dict]]) -> dict[str, tuple]:
    """Return the legacy tuple map used by the forecast for MI and discount stock."""

    all_skus = {sku for offers in price_maps.values() for sku in offers}
    result: dict[str, tuple] = {}
    for sku in all_skus:
        choices = [offers[sku] for offers in price_maps.values() if sku in offers and offers[sku]["in_stock"]]
        if not choices:
            choices = [offers[sku] for offers in price_maps.values() if sku in offers]
        if not choices:
            continue
        best = min(choices, key=lambda item: item["price"])
        max_discount = max((item.get("discount_pct", 0) for item in choices), default=0)
        result[sku] = (best["price"], any(item["in_stock"] for item in choices), max_discount)
    return result


def _landed_price(offer: dict, config: SupplierConfig) -> float:
    return round(float(offer["price"]) * (1 + config.price_adjustment_pct / 100), 6)


def _margin_pct(sell_price: float | None, landed_price: float) -> float | None:
    if not sell_price or sell_price <= 0:
        return None
    return round((sell_price - landed_price) / sell_price * 100, 2)


def _rank_candidates(candidates: list[dict], rules: dict) -> list[dict]:
    if not candidates:
        return []
    strategy = rules.get("strategy", "priority_within_tolerance")
    if strategy == "lowest_price":
        return sorted(candidates, key=lambda item: (item["landed_price"], item["lead_time_days"], item["priority"]))

    if strategy == "balanced":
        min_price = min(item["landed_price"] for item in candidates)
        max_lead = max(max(item["lead_time_days"] for item in candidates), 1)
        price_weight = float(rules.get("price_weight", 0.7))
        lead_weight = float(rules.get("lead_time_weight", 0.2))
        priority_weight = max(0.0, 1 - price_weight - lead_weight)
        for item in candidates:
            item["score"] = (
                price_weight * item["landed_price"] / min_price
                + lead_weight * item["lead_time_days"] / max_lead
                + priority_weight * item["priority"]
            )
        return sorted(candidates, key=lambda item: (item["score"], item["landed_price"]))

    # Prefer the configured supplier while its landed cost stays within the tolerance.
    tolerance = max(0.0, float(rules.get("price_tolerance_pct", 5)))
    cheapest = min(item["landed_price"] for item in candidates)
    ceiling = cheapest * (1 + tolerance / 100)
    return sorted(
        candidates,
        key=lambda item: (
            0 if item["landed_price"] <= ceiling else 1,
            item["priority"] if item["landed_price"] <= ceiling else item["landed_price"],
            item["landed_price"],
            item["lead_time_days"],
        ),
    )


def _allocate_once(
    demand_lines: list[dict],
    price_maps: dict[str, dict[str, dict]],
    configs: dict[str, SupplierConfig],
    rules: dict,
    active_suppliers: set[str],
) -> dict:
    orders = {name: [] for name in active_suppliers}
    unallocated: list[dict] = []
    supplier_spend = {name: 0.0 for name in active_suppliers}
    used_qty: dict[tuple[str, str], float] = {}
    max_lead = float(rules.get("max_lead_time_days", 0) or 0)
    min_margin = float(rules.get("min_purchase_margin_pct", 0) or 0)

    for line in sorted(demand_lines, key=lambda item: (item.get("days_left", 999), clean_text(item.get("sku")))):
        sku = clean_text(line.get("sku"))
        requested = max(0.0, float(line.get("order_qty", 0) or 0))
        remaining = requested
        if not sku or requested <= 0:
            continue

        candidates: list[dict] = []
        rejection_reasons: set[str] = set()
        for supplier in active_suppliers:
            config = configs[supplier]
            offer = price_maps.get(supplier, {}).get(sku)
            if not offer:
                rejection_reasons.add("SKU відсутній у прайсі")
                continue
            if not offer.get("in_stock", False):
                rejection_reasons.add("немає в наявності")
                continue
            landed = _landed_price(offer, config)
            lead_time = float(offer.get("lead_time_days") or config.lead_time_days)
            margin = _margin_pct(line.get("sell_price"), landed)
            moq = max(float(offer.get("min_order_qty") or 0), config.min_order_qty)
            if requested < moq:
                rejection_reasons.add(f"кількість менша MOQ ({moq:g})")
                continue
            if max_lead and lead_time > max_lead:
                rejection_reasons.add(f"lead time понад {max_lead:g} дн")
                continue
            if margin is not None and margin < min_margin:
                rejection_reasons.add(f"маржа нижча {min_margin:g}%")
                continue
            available = offer.get("available_qty")
            if available is not None:
                available = max(0.0, float(available) - used_qty.get((supplier, sku), 0))
                if available <= 0:
                    rejection_reasons.add("вичерпано доступну кількість")
                    continue
            candidates.append(
                {
                    "supplier": supplier,
                    "offer": offer,
                    "landed_price": landed,
                    "lead_time_days": lead_time,
                    "margin_pct": margin,
                    "priority": config.priority,
                    "available_qty": available,
                    "min_order_qty": moq,
                }
            )

        for candidate in _rank_candidates(candidates, rules):
            if remaining <= 0:
                break
            supplier = candidate["supplier"]
            config = configs[supplier]
            max_qty = remaining
            if candidate["available_qty"] is not None:
                max_qty = min(max_qty, candidate["available_qty"])
            if config.order_limit_value:
                budget_left = config.order_limit_value - supplier_spend[supplier]
                max_qty = min(max_qty, math.floor(max(budget_left, 0) / candidate["landed_price"] + 1e-9))
            qty = math.floor(max_qty + 1e-9)
            if qty <= 0:
                rejection_reasons.add("вичерпано ліміт постачальника")
                continue
            if qty < candidate["min_order_qty"]:
                rejection_reasons.add(f"залишок кількості менший MOQ ({candidate['min_order_qty']:g})")
                continue
            value = round(qty * candidate["landed_price"], 2)
            orders[supplier].append(
                {
                    "supplier": supplier,
                    "sku": sku,
                    "supplier_sku": candidate["offer"].get("supplier_sku", sku),
                    "barcode": candidate["offer"].get("barcode", ""),
                    "name": clean_text(line.get("name")),
                    "quantity": qty,
                    "unit_price": round(float(candidate["offer"]["price"]), 4),
                    "landed_unit_price": round(candidate["landed_price"], 4),
                    "order_value": value,
                    "discount_pct": candidate["offer"].get("discount_pct", 0),
                    "lead_time_days": candidate["lead_time_days"],
                    "margin_pct": candidate["margin_pct"],
                    "days_left": line.get("days_left"),
                }
            )
            supplier_spend[supplier] += value
            used_qty[(supplier, sku)] = used_qty.get((supplier, sku), 0) + qty
            remaining -= qty

        if remaining > 0:
            unallocated.append(
                {
                    "sku": sku,
                    "name": clean_text(line.get("name")),
                    "requested_qty": requested,
                    "unallocated_qty": remaining,
                    "days_left": line.get("days_left"),
                    "reason": "; ".join(sorted(rejection_reasons)) or "немає придатної пропозиції",
                }
            )

    return {"orders": orders, "unallocated": unallocated, "supplier_spend": supplier_spend}


def allocate_orders(
    demand_lines: list[dict],
    price_maps: dict[str, dict[str, dict]],
    supplier_configs: list[dict],
    rules: dict | None = None,
) -> dict:
    """Allocate each SKU, split when necessary, and enforce supplier order minima.

    Supplier-level minimum order values are resolved by rerunning allocation without
    suppliers that fail their minimum. This keeps budgets and stock quantities exact.
    """

    rules = dict(rules or {})
    configs = {cfg.name: cfg for cfg in (SupplierConfig.from_mapping(item) for item in supplier_configs)}
    active = {name for name, config in configs.items() if config.enabled and name in price_maps}
    if not active:
        return {
            "orders": {},
            "unallocated": [
                {
                    "sku": clean_text(line.get("sku")),
                    "name": clean_text(line.get("name")),
                    "requested_qty": line.get("order_qty", 0),
                    "unallocated_qty": line.get("order_qty", 0),
                    "days_left": line.get("days_left"),
                    "reason": "не завантажено жодного активного прайсу",
                }
                for line in demand_lines
                if float(line.get("order_qty", 0) or 0) > 0
            ],
            "summary": {},
            "excluded_suppliers": [],
        }

    excluded: list[dict] = []
    while active:
        result = _allocate_once(demand_lines, price_maps, configs, rules, active)
        below_minimum = [
            name
            for name in active
            if result["orders"].get(name)
            and configs[name].min_order_value > result["supplier_spend"].get(name, 0)
        ]
        if not below_minimum:
            break
        for name in below_minimum:
            excluded.append(
                {
                    "supplier": name,
                    "order_value": round(result["supplier_spend"].get(name, 0), 2),
                    "min_order_value": configs[name].min_order_value,
                }
            )
            active.remove(name)
    else:
        result = _allocate_once(demand_lines, price_maps, configs, rules, set())

    summary = {}
    for name, lines in result["orders"].items():
        if not lines:
            continue
        summary[name] = {
            "sku_count": len(lines),
            "quantity": sum(line["quantity"] for line in lines),
            "order_value": round(sum(line["order_value"] for line in lines), 2),
        }
    result["summary"] = summary
    result["excluded_suppliers"] = excluded
    return result
