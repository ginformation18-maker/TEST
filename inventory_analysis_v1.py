#!/usr/bin/env python3
"""ØªØ­ÙÙÙ ÙÙØ·Ø© Ø£Ø¹ÙØ§Ø± Ø§ÙÙØ®Ø²ÙÙ Ø¨ÙØºØ© Python 3.12.

Ø§ÙØªØ«Ø¨ÙØª: python -m pip install "pandas>=2.2,<3" "XlsxWriter>=3.2,<4"
Ø§ÙØªØ´ØºÙÙ: python inventory_analysis.py inventory_20261114.csv -o report.xlsx
Ø§ÙÙØ³Ø§Ø¹Ø¯Ø©: python inventory_analysis.py --help

ØªÙÙØ±Ø£ Ø´Ø±Ø§Ø¦Ø­ Ø§ÙØ£Ø¹ÙØ§Ø± ÙÙ Ø§ÙÙÙØ·Ø©. ØªØ§Ø±ÙØ® Ø§ÙÙÙØ·Ø© ÙØ¤Ø«Ø± ÙÙ Ø­Ø³Ø§Ø¨ Ø£ÙØ§Ù Ø§ÙØ­Ø±ÙØ©.
Ø´ØºÙÙ Ø§ÙØ¯Ø§ÙØ© analyze_inventory(InventoryConfig(...)) ÙÙØ§Ø³ØªØ¹ÙØ§Ù ÙÙ ÙÙØ¯ Ø¢Ø®Ø±.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter

import numpy as np
import pandas as pd
from xlsxwriter import Workbook
from xlsxwriter.exceptions import XlsxWriterException
from xlsxwriter.worksheet import Worksheet


SKU, NAME, CATEGORY, SUPPLIER = "Ø±ÙÙ Ø§ÙØµÙÙ", "Ø§Ø³Ù Ø§ÙØµÙÙ", "Ø§ÙÙØ¦Ø©", "Ø§Ø³Ù Ø§ÙÙÙØ±Ø¯"
QTY, VALUE = "Ø§ÙØ±ØµÙØ¯", "Ø§ÙÙÙÙØ©"
SALE, RECEIPT = "Ø¢Ø®Ø± Ø¨ÙØ¹", "Ø¢Ø®Ø± ØªÙØ±ÙØ¯"
UNIT, DECLARED_COST = "ÙØ­Ø¯Ø© Ø§ÙÙÙØ§Ø³", "ØªÙÙÙØ© Ø§ÙÙØ­Ø¯Ø© Ø§ÙÙØµØ±ÙØ­ Ø¨ÙØ§"
IMPLIED_COST, MARGIN = "ØªÙÙÙØ© Ø§ÙÙØ­Ø¯Ø© Ø§ÙÙØ­Ø³ÙØ¨Ø©", "ÙØ§ÙØ´ Ø§ÙØ±Ø¨Ø­"
OLD_VALUE, FROZEN_VALUE = "ÙÙÙØ© Ø£ÙØ«Ø± ÙÙ 180 ÙÙÙÙØ§", "Ø§ÙÙÙÙØ© Ø§ÙÙØ¬ÙØ¯Ø©"
LAST_EVENT, DAYS_IDLE = "Ø¢Ø®Ø± ØªØ§Ø±ÙØ® ÙØªØ§Ø­", "Ø£ÙØ§Ù ÙÙØ° Ø¢Ø®Ø± ØªØ§Ø±ÙØ®"
UNIT_UNKNOWN = "ÙØ­Ø¯Ø© ØºÙØ± ÙÙØ«ÙØ©"
FROZEN_IDLE_DAYS = 180
EXCEL_MAX_ROWS = 1_048_576

PERIODS = pd.Index([
    "0-15 ÙÙÙ", "16-30 ÙÙÙ", "31-45 ÙÙÙ", "46-60 ÙÙÙ", "61-75 ÙÙÙ", "76-90 ÙÙÙ",
    "91-105 ÙÙÙ", "106-120 ÙÙÙ", "121-135 ÙÙÙ", "136-150 ÙÙÙ", "151-165 ÙÙÙ",
    "166-180 ÙÙÙ", "181 ÙÙÙ ÙØ£ÙØ«Ø±",
])
AGE_HEADERS = PERIODS.repeat(2) + np.tile([" ÙÙÙØ©", " ÙÙÙØ©"], len(PERIODS))
AGES = ("Ø­ØªÙ 90 ÙÙÙÙØ§", "91â180 ÙÙÙÙØ§", "Ø£ÙØ«Ø± ÙÙ 180 ÙÙÙÙØ§", "ØºÙØ± ÙÙØ²ÙØ¹ Ø¹ÙØ±ÙÙØ§")
AGE_DETAIL = (
    "ÙÙÙØ© Ø­ØªÙ 90 ÙÙÙÙØ§", "ÙÙÙØ© Ø­ØªÙ 90 ÙÙÙÙØ§", "ÙÙÙØ© 91â180 ÙÙÙÙØ§",
    "ÙÙÙØ© 91â180 ÙÙÙÙØ§", "ÙÙÙØ© Ø£ÙØ«Ø± ÙÙ 180 ÙÙÙÙØ§", OLD_VALUE,
)
NUMBER_TRANSLATION = str.maketrans(
    "Ù Ù¡Ù¢Ù£Ù¤Ù¥Ù¦Ù§Ù¨Ù©Û°Û±Û²Û³Û´ÛµÛ¶Û·Û¸Û¹Ù«Ù¬â", "01234567890123456789.,-"
)
MISSING = ("", "-", "â", "N/A", "NULL", "null", "ØºÙØ± ÙØªÙÙØ±", "ØºÙØ± ÙØªØ§Ø­")


@dataclass(frozen=True, slots=True)
class InventoryConfig:
    source: Path
    output: Path
    as_of: date | None = None
    encoding: str = "auto"
    separator: str = ","
    date_format: str = "%Y-%m-%d"
    uom_column: str | None = None
    unit_cost_column: str | None = None
    margin_column: str | None = None
    last_movement_column: str | None = None
    margin_is_ratio: bool = False
    cost_tolerance: float = 0.01
    value_tolerance: float = 0.01
    quantity_tolerance: float = 0.001
    relative_cost_tolerance: float = 0.01
    abc_a: float = 0.80
    abc_b: float = 0.95
    verbose: bool = True


@contextmanager
def timed(label: str, verbose: bool):
    started = perf_counter()
    try:
        yield
    finally:
        if verbose:
            print(f"{label}  {perf_counter() - started:.3f} Ø«", flush=True)


def headers(index: pd.Index) -> pd.Index:
    return (
        index.str.normalize("NFKC")
        .str.translate(NUMBER_TRANSLATION)
        .str.replace(r"[\u061c\u200e\u200f\ufeff]", "", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"\s*[-ââââ]\s*", "-", regex=True)
        .str.strip()
    )


def clean(series: pd.Series) -> pd.Series:
    text = series.str.strip()
    return text.mask(text.isin(MISSING) | text.eq(""))


def numbers(series: pd.Series) -> pd.Series:
    try:
        parsed = series.to_numpy(dtype="float64", na_value=np.nan)
    except (TypeError, ValueError):
        text = (
            series.str.strip().str.translate(NUMBER_TRANSLATION)
            .str.replace(r"^\(\s*([^()]*)\s*\)$", r"-\1", regex=True)
        )
        grouped = text.str.fullmatch(
            r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?(?:[eE][+-]?\d+)?", na=False
        )
        text = text.mask(grouped, text.str.replace(",", "", regex=False))
        parsed = pd.to_numeric(text, errors="coerce").to_numpy(dtype="float64", na_value=np.nan)
    return pd.Series(np.where(np.isfinite(parsed), parsed, np.nan), index=series.index)


def snapshot_date(config: InventoryConfig) -> pd.Timestamp:
    if config.as_of is not None:
        return pd.Timestamp(config.as_of)
    found = re.search(r"(?<!\d)(\d{8})(?!\d)", config.source.stem)
    if found:
        try:
            return pd.Timestamp(datetime.strptime(found[1], "%Y%m%d").date())
        except ValueError:
            pass
    raise ValueError("Ø­Ø¯ÙØ¯ ØªØ§Ø±ÙØ® ÙÙØ·Ø© Ø§ÙÙØ®Ø²ÙÙ Ø¨Ø§Ø³ØªØ®Ø¯Ø§Ù --as-of YYYY-MM-DD.")


def read_inventory(config: InventoryConfig) -> tuple[pd.DataFrame, str]:
    options = {
        "sep": config.separator, "dtype": "string", "keep_default_na": False,
        "engine": "c", "encoding_errors": "strict", "on_bad_lines": "error",
    }
    encoding = "utf-8-sig" if config.encoding == "auto" else config.encoding
    try:
        raw = pd.read_csv(config.source, encoding=encoding, **options)
    except UnicodeDecodeError:
        if config.encoding != "auto":
            raise
        encoding = "cp1256"
        raw = pd.read_csv(config.source, encoding=encoding, **options)
    with config.source.open(encoding=encoding, newline="") as source:
        source_headers = next(filter(None, csv.reader(source, delimiter=config.separator)), ())
    if not isinstance(raw.index, pd.RangeIndex) or len(source_headers) != len(raw.columns):
        raise ValueError("Ø¹Ø¯Ø¯ Ø§ÙØ­ÙÙÙ ÙÙ Ø³Ø¬ÙØ§Øª CSV ÙØ§ ÙØ·Ø§Ø¨Ù ØµÙ Ø§ÙØ¹ÙØ§ÙÙÙ.")
    raw.columns = headers(pd.Index(source_headers, dtype="str"))
    if not raw.columns.is_unique:
        raise ValueError("Ø¹ÙØ§ÙÙÙ CSV ÙØªÙØ±Ø±Ø© Ø¨Ø¹Ø¯ ØªÙØ­ÙØ¯ Ø§ÙÙØ³Ø§ÙØ§Øª ÙØ§ÙØ´Ø±Ø·Ø§Øª.")
    required = pd.Index([SKU, NAME, QTY, VALUE]).append(AGE_HEADERS)
    missing = required.difference(raw.columns, sort=False)
    if not missing.empty:
        raise ValueError("Ø£Ø¹ÙØ¯Ø© ÙØ§Ø²ÙØ© ÙÙÙÙØ¯Ø© ÙÙ CSV: " + "Ø ".join(missing))
    return raw, encoding


def optional_column(
    raw: pd.DataFrame, override: str | None, aliases: tuple[str, ...]
) -> tuple[pd.Series, str | None]:
    if override:
        column = headers(pd.Index([override]))[0]
        if column not in raw:
            raise ValueError(f"Ø§ÙØ¹ÙÙØ¯ Ø§ÙÙØ­Ø¯Ø¯ ØºÙØ± ÙÙØ¬ÙØ¯ ÙÙ CSV: {override}")
        return clean(raw[column]), column
    candidates = pd.Index(aliases).intersection(raw.columns, sort=False)
    if candidates.empty:
        return pd.Series(pd.NA, index=raw.index, dtype="string"), None
    column = candidates[0]
    return clean(raw[column]), column


@dataclass(frozen=True, slots=True)
class Sources:
    unit: str | None
    cost: str | None
    margin: str | None
    movement: str | None
    sale: str | None
    receipt: str | None


def prepare_inventory(raw: pd.DataFrame, config: InventoryConfig, as_of: pd.Timestamp):
    unit_text, unit_col = optional_column(raw, config.uom_column, (UNIT, "Ø§ÙÙØ­Ø¯Ø©", "UoM", "UOM"))
    cost_text, cost_col = optional_column(raw, config.unit_cost_column, (DECLARED_COST, "ØªÙÙÙØ© Ø§ÙÙØ­Ø¯Ø©", "Unit Cost"))
    margin_text, margin_col = optional_column(raw, config.margin_column, (MARGIN, "ÙØ³Ø¨Ø© ÙØ§ÙØ´ Ø§ÙØ±Ø¨Ø­", "Gross Margin"))
    move_text, move_col = optional_column(
        raw, config.last_movement_column, ("Ø¢Ø®Ø± Ø­Ø±ÙØ©", "ØªØ§Ø±ÙØ® Ø¢Ø®Ø± Ø­Ø±ÙØ©", "Last Movement")
    )
    sale_text, sale_col = optional_column(raw, None, (SALE, "ØªØ§Ø±ÙØ® Ø¢Ø®Ø± Ø¨ÙØ¹"))
    receipt_text, receipt_col = optional_column(raw, None, (RECEIPT, "ØªØ§Ø±ÙØ® Ø¢Ø®Ø± ØªÙØ±ÙØ¯"))
    category, _ = optional_column(raw, None, (CATEGORY,))
    supplier, _ = optional_column(raw, None, (SUPPLIER,))
    sources = Sources(unit_col, cost_col, margin_col, move_col, sale_col, receipt_col)

    # ØªÙØ­ÙØ¯ ØªÙØ«ÙÙ ÙØ¹Ø±ÙÙØ§Øª CSV ÙØ¹ Ø­ÙØ¸ Ø§ÙØ£ØµÙØ§Ø± Ø§ÙØ¨Ø§Ø¯Ø¦Ø©.
    sku = clean(raw[SKU].str.replace(r'^="(.*)"$', r"\1", regex=True))
    qty, value = numbers(raw[QTY]), numbers(raw[VALUE])
    positive, negative = qty.gt(0), qty.lt(0)
    active = positive | negative
    unit = unit_text.str.normalize("NFKC").str.replace(r"\s+", " ", regex=True).str.casefold()
    unit_count = unit.groupby(sku, sort=False).transform("nunique")
    mixed_unit = unit_count.gt(1)
    partial_unit = unit_count.eq(1) & unit.isna().groupby(sku, sort=False).transform("any").eq(True)
    unit_basis = unit.fillna(UNIT_UNKNOWN).mask(mixed_unit, "ÙØ­Ø¯Ø§Øª ÙØªØ¹Ø§Ø±Ø¶Ø©").mask(
        partial_unit, "ØªÙØ«ÙÙ Ø¬Ø²Ø¦Ù ÙÙÙØ­Ø¯Ø©"
    )

    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        implied = value.div(qty.where(active))
    implied = implied.where(np.isfinite(implied))
    declared = numbers(cost_text)
    bad_implied = active & (implied.isna() | implied.le(0))
    bad_declared = cost_text.notna() & (declared.isna() | declared.le(0))
    missing_declared = active & (cost_col is not None) & cost_text.isna()
    difference = active & implied.gt(0) & declared.gt(0) & ~np.isclose(
        implied, declared, atol=config.cost_tolerance, rtol=config.relative_cost_tolerance
    )
    comparable = implied.where(implied.gt(0) & unit.notna() & ~mixed_unit & ~partial_unit)
    same_sku = comparable.groupby(sku, sort=False)
    low_cost, high_cost = same_sku.transform("min"), same_sku.transform("max")
    spread = comparable.notna() & high_cost.sub(low_cost).gt(
        config.cost_tolerance + config.relative_cost_tolerance * low_cost.abs()
    )
    cost_check = np.select(
        [bad_implied, bad_declared, difference, spread, missing_declared, active & declared.gt(0), active],
        ["ØªÙÙÙØ© ÙØ­Ø³ÙØ¨Ø© ØºÙØ± ØµØ§ÙØ­Ø©", "ØªÙÙÙØ© ÙØµØ±ÙØ­ Ø¨ÙØ§ ØºÙØ± ØµØ§ÙØ­Ø©", "ÙØ±Ù Ø¹Ù Ø§ÙØªÙÙÙØ© Ø§ÙÙØµØ±ÙØ­ Ø¨ÙØ§",
         "ØªÙØ§ÙØª ØªÙÙÙØ© ÙÙÙØ³ Ø§ÙØµÙÙ ÙØ§ÙÙØ­Ø¯Ø©", "ØªÙÙÙØ© ÙØµØ±ÙØ­ Ø¨ÙØ§ ÙÙÙÙØ¯Ø©", "ÙØ·Ø§Ø¨ÙØ© ÙÙÙØµØ±ÙØ­ Ø¨ÙØ§", "ØªÙÙÙØ© ÙØ­Ø³ÙØ¨Ø©"],
        default="",
    )

    margin_text = margin_text.str.translate(NUMBER_TRANSLATION).str.replace("Ùª", "%", regex=False)
    margin = numbers(margin_text.str.removesuffix("%"))
    margin = margin.div(np.where(margin_text.str.endswith("%", na=False) | (not config.margin_is_ratio), 100, 1))
    margin_bad = margin_text.notna() & margin.isna()
    margin_loss = positive & margin.lt(0)
    parsed_dates = pd.to_datetime(
        pd.concat([sale_text, receipt_text, move_text], ignore_index=True),
        format=config.date_format, errors="coerce",
    ).dt.normalize()
    dates = pd.DataFrame(
        parsed_dates.to_numpy().reshape((len(raw), 3), order="F"),
        index=raw.index, columns=["sale", "receipt", "movement"],
    )
    valid_dates = dates.where(dates.le(as_of))
    date_bad = (
        (sale_text.notna() & valid_dates["sale"].isna())
        | (receipt_text.notna() & valid_dates["receipt"].isna())
        | (move_text.notna() & valid_dates["movement"].isna())
    )
    last_observed = valid_dates[["sale", "receipt"]].max(axis=1)
    direct = valid_dates["movement"]
    movement_conflict = direct.lt(last_observed)
    direct_ok = direct.notna() & ~movement_conflict
    both = valid_dates[["sale", "receipt"]].notna().all(axis=1)
    latest = direct.where(direct_ok, last_observed)
    idle = (as_of - latest).dt.days
    proxy_stale = ~direct_ok & both & idle.gt(FROZEN_IDLE_DAYS)
    direct_stale = direct_ok & idle.gt(FROZEN_IDLE_DAYS)
    activity_unknown = ~(direct_ok | both | idle.le(FROZEN_IDLE_DAYS))
    activity_basis = np.select(
        [direct_ok, both, latest.notna()],
        ["ØªØ§Ø±ÙØ® Ø§ÙØ­Ø±ÙØ©", "Ø§ÙØ¨ÙØ¹ ÙØ§ÙØªÙØ±ÙØ¯", "ØªØ§Ø±ÙØ® Ø¬Ø²Ø¦Ù"], default="ØºÙØ± ÙØªØ§Ø­",
    )

    detail = pd.DataFrame({
        "record": np.arange(1, len(raw) + 1), "sku": sku, "name": clean(raw[NAME]),
        "category": category, "supplier": supplier, "qty": qty, "value": value,
        "positive": positive, "negative": negative, "unit": unit,
        "unit_basis": unit_basis, "qty_rankable": ~(mixed_unit | partial_unit),
        "declared": declared, "implied": implied, "cost_check": cost_check, "margin": margin,
        "margin_loss": margin_loss, "last_event": latest, "idle_days": idle,
        "activity_basis": activity_basis, "direct_stale": direct_stale,
        "proxy_stale": proxy_stale, "activity_unknown": activity_unknown,
    })
    flags = np.column_stack([
        negative, sku.isna(), qty.isna() | value.isna(), qty.eq(0) & value.notna() & value.ne(0),
        bad_implied, bad_declared, missing_declared, difference, spread, mixed_unit,
        active & (unit_col is not None) & unit.isna(), date_bad, movement_conflict, margin_bad, margin_loss,
    ])
    labels = np.array([
        "Ø±ØµÙØ¯ Ø³Ø§ÙØ¨Ø ", "Ø±ÙÙ Ø§ÙØµÙÙ ÙÙÙÙØ¯Ø ", "Ø±ØµÙØ¯ Ø£Ù ÙÙÙØ© ØºÙØ± ØµØ§ÙØ­ÙÙØ ", "ÙÙÙØ© ÙØ±ØµÙØ¯ ØµÙØ±ÙØ ",
        "ØªÙÙÙØ© ÙØ­Ø³ÙØ¨Ø© ØµÙØ±ÙØ© Ø£Ù Ø³Ø§ÙØ¨Ø© Ø£Ù ÙÙÙÙØ¯Ø©Ø ", "ØªÙÙÙØ© ÙØµØ±ÙØ­ Ø¨ÙØ§ ØºÙØ± ØµØ§ÙØ­Ø©Ø ",
        "ØªÙÙÙØ© ÙØµØ±ÙØ­ Ø¨ÙØ§ ÙÙÙÙØ¯Ø©Ø ", "ÙØ±Ù ØªÙÙÙØ© Ø§ÙÙØ­Ø¯Ø©Ø ", "ØªÙØ§ÙØª ØªÙÙÙØ© ÙÙÙØ³ Ø§ÙØµÙÙ ÙØ§ÙÙØ­Ø¯Ø©Ø ", "ØªØ¹Ø§Ø±Ø¶ ÙØ­Ø¯Ø§Øª Ø§ÙØµÙÙØ ",
        "ÙØ­Ø¯Ø© ÙÙÙÙØ¯Ø©Ø ", "ØªØ§Ø±ÙØ® ØºÙØ± ØµØ§ÙØ­ Ø£Ù Ø¨Ø¹Ø¯ Ø§ÙÙÙØ·Ø©Ø ", "ØªØ§Ø±ÙØ® Ø§ÙØ­Ø±ÙØ© Ø£ÙØ¯Ù ÙÙ Ø¨ÙØ¹ Ø£Ù ØªÙØ±ÙØ¯ ÙØ³Ø¬ÙØ ",
        "ÙØ§ÙØ´ ØºÙØ± ØµØ§ÙØ­Ø ", "ÙØ§ÙØ´ Ø³Ø§ÙØ¨Ø ",
    ], dtype=object)
    reasons = np.where(flags, labels, "")
    detail["reasons"] = pd.Series(reasons[:, 0], index=detail.index, dtype="string").str.cat(reasons[:, 1:])
    return detail, sources


def analyze_aging(detail: pd.DataFrame, raw: pd.DataFrame, config: InventoryConfig) -> pd.DataFrame:
    flat = pd.Series(raw.loc[:, AGE_HEADERS].to_numpy().ravel(), dtype="string")
    amounts = numbers(flat).to_numpy().reshape((len(raw), len(PERIODS), 2))
    young = amounts[:, :6, :].sum(axis=1)
    middle = amounts[:, 6:12, :].sum(axis=1)
    old = amounts[:, 12, :]
    totals = amounts.sum(axis=1)
    reconciled = (
        np.isfinite(amounts).all(axis=(1, 2))
        & np.isclose(totals[:, 0], detail["qty"], atol=config.quantity_tolerance, rtol=1e-10)
        & np.isclose(totals[:, 1], detail["value"], atol=config.value_tolerance, rtol=1e-10)
    )
    positive = detail["positive"]
    incoherent = positive & (
        (amounts < 0).any(axis=(1, 2))
        | ((amounts[:, :, 0] == 0) & (amounts[:, :, 1] > 0)).any(axis=1)
    )
    valid = reconciled & ~incoherent
    detail["reasons"] += np.where(~valid, "Ø´Ø±Ø§Ø¦Ø­ Ø§ÙØ£Ø¹ÙØ§Ø± ÙØ§ ØªØ·Ø§Ø¨Ù Ø§ÙØ±ØµÙØ¯ Ø£Ù Ø§ÙÙÙÙØ©Ø ", "")
    detail["reasons"] = detail["reasons"].str.removesuffix("Ø ")

    aging = pd.DataFrame(
        np.column_stack([young, middle, old]), index=detail.index, columns=AGE_DETAIL
    ).where(valid & positive, axis=0)
    aging["age_valid"] = valid
    stocked_old = positive & valid & aging[AGE_DETAIL[4]].gt(0)
    aging["unknown_old"] = stocked_old & detail["activity_unknown"]
    aging["stalled_direct"] = stocked_old & detail["direct_stale"]
    aging["stalled_proxy"] = stocked_old & detail["proxy_stale"]
    stalled = aging["stalled_direct"] | aging["stalled_proxy"]
    loss = detail["margin_loss"]
    aging["stall_value"] = aging[OLD_VALUE].where(stalled, 0)
    aging["loss_value"] = detail["value"].where(loss & detail["value"].ge(0), 0)
    # Ø§ÙÙØ§ÙØ´ Ø§ÙØ³Ø§ÙØ¨ ÙØ´ÙÙ ÙÙÙØ© Ø§ÙØ±ØµÙØ¯ Ø§ÙÙÙØ¬Ø¨ ÙØ§ÙÙØ©. ØªÙØ§Ø·Ø¹ Ø§ÙØ´Ø±Ø·ÙÙ ÙÙØ­Ø³Ø¨ ÙØ±Ø© ÙØ§Ø­Ø¯Ø©.
    aging[FROZEN_VALUE] = aging["stall_value"].where(~loss, detail["value"].where(detail["value"].ge(0)))
    aging["freeze_reason"] = np.select(
        [stalled & loss, loss, aging["stalled_direct"], aging["stalled_proxy"], aging["unknown_old"]],
        ["Ø±ÙÙØ¯ ÙÙØ§ÙØ´ Ø³Ø§ÙØ¨", "ÙØ§ÙØ´ Ø³Ø§ÙØ¨", "Ø±ÙÙØ¯ Ø­Ø³Ø¨ Ø§ÙØ­Ø±ÙØ©", "Ø±ÙÙØ¯ Ø­Ø³Ø¨ Ø§ÙØ¨ÙØ¹ ÙØ§ÙØªÙØ±ÙØ¯",
         "Ø£Ø¯ÙØ© Ø§ÙØ­Ø±ÙØ© ØºÙØ± ÙØ§ÙÙØ©"], default="",
    )
    return aging


def pareto(items: pd.DataFrame, metric: str, scope: str, config: InventoryConfig) -> pd.DataFrame:
    ranked = (
        items.loc[items[metric].gt(0), ["sku", scope, metric]]
        .sort_values([scope, metric, "sku"], ascending=[True, False, True])
        .set_index("sku")
    )
    grouped = ranked.groupby(scope, sort=False)[metric]
    total, cumulative = grouped.transform("sum"), grouped.cumsum()
    previous = cumulative.sub(ranked[metric]).div(total)
    return ranked.assign(
        rank=grouped.cumcount() + 1,
        cumulative=cumulative.div(total).clip(0, 1),
        abc=np.select(
            [previous.lt(config.abc_a - 1e-12), previous.lt(config.abc_b - 1e-12)],
            ["A", "B"], default="C",
        ),
    )


def classify_inventory(detail: pd.DataFrame, config: InventoryConfig):
    active = detail.loc[detail["positive"] & detail["sku"].notna()].copy()
    active["valid_value"] = active["value"].where(active["value"].gt(0))
    items = active.groupby("sku", sort=False).agg(
        name=("name", "first"), value=("valid_value", "sum"), qty=("qty", "sum"),
        qty_rankable=("qty_rankable", "all"), unit_basis=("unit_basis", "first"),
    ).reset_index()
    items["qty"] = items["qty"].where(items["qty_rankable"])
    items["all_inventory"] = "Ø¬ÙÙØ¹ Ø§ÙØ£ØµÙØ§Ù"
    by_value = pareto(items, "value", "all_inventory", config)
    by_quantity = pareto(items, "qty", "unit_basis", config)
    cols = ["rank", "cumulative", "abc"]
    report = (
        items.set_index("sku").join(by_value[cols].add_prefix("value_"))
        .join(by_quantity[cols].add_prefix("qty_"))
        .sort_values("value_rank", kind="stable").reset_index()
    )
    return report, by_value, by_quantity


AUDIT_LABELS = {
    "sku": SKU, "name": NAME, "category": CATEGORY, "supplier": SUPPLIER, "qty": QTY,
    "value": VALUE, "unit_basis": "Ø­Ø§ÙØ© ÙØ­Ø¯Ø© Ø§ÙÙÙØ§Ø³", "implied": IMPLIED_COST,
    "declared": DECLARED_COST, "cost_check": "ØªØ¯ÙÙÙ ØªÙÙÙØ© Ø§ÙÙØ­Ø¯Ø©", "margin": MARGIN,
    "last_event": LAST_EVENT, "reasons": "Ø£Ø³Ø¨Ø§Ø¨ Ø§ÙØªØ¯ÙÙÙ", "record": "Ø±ÙÙ Ø³Ø¬Ù CSV",
}
AGING_LABELS = {
    "sku": SKU, "name": NAME, "category": CATEGORY, "qty": QTY, "value": VALUE,
    "last_event": LAST_EVENT, "idle_days": DAYS_IDLE, "activity_basis": "Ø£Ø³Ø§Ø³ ØªØ§Ø±ÙØ® Ø§ÙØ­Ø±ÙØ©",
    "margin": MARGIN, "freeze_reason": "Ø­Ø§ÙØ© Ø§ÙØªØ¬ÙÙØ¯", "age_status": "ÙØ·Ø§Ø¨ÙØ© Ø´Ø±Ø§Ø¦Ø­ Ø§ÙØ£Ø¹ÙØ§Ø±",
}
ABC_LABELS = {
    "sku": SKU, "name": NAME, "value": "Ø§ÙÙÙÙØ© Ø§ÙÙØ¤ÙÙØ©", "value_rank": "ØªØ±ØªÙØ¨ Ø§ÙÙÙÙØ©",
    "value_cumulative": "Ø§ÙÙÙÙØ© Ø§ÙØªØ±Ø§ÙÙÙØ© %", "value_abc": "ABC Ø§ÙÙÙÙØ©",
    "qty": "Ø§ÙÙÙÙØ© Ø§ÙÙØ¤ÙÙØ©", "unit_basis": "ÙØ­Ø¯Ø© ÙÙØ§Ø±ÙØ© Ø§ÙÙÙÙØ©", "qty_rank": "ØªØ±ØªÙØ¨ Ø§ÙÙÙÙØ©",
    "qty_cumulative": "Ø§ÙÙÙÙØ© Ø§ÙØªØ±Ø§ÙÙÙØ© %", "qty_abc": "ABC Ø§ÙÙÙÙØ©",
}


def report_tables(detail: pd.DataFrame, aging: pd.DataFrame, abc: pd.DataFrame, sources: Sources):
    audit = detail.loc[detail["reasons"].ne("")].copy()
    audit["abs_value"] = audit["value"].abs()
    audit = (
        audit.sort_values(["negative", "abs_value", "record"], ascending=[False, False, True])
        .loc[:, list(AUDIT_LABELS)].rename(columns=AUDIT_LABELS)
    )
    positive = detail["positive"]
    age = detail.loc[positive, ["sku", "name", "category", "qty", "value"]].join(
        aging.loc[:, [*AGE_DETAIL, FROZEN_VALUE, "freeze_reason"]]
    )
    age = age.join(detail[["last_event", "idle_days", "activity_basis", "margin"]])
    age["age_status"] = np.where(aging.loc[age.index, "age_valid"], "ÙØ·Ø§Ø¨ÙØ©", "ØªØ­Øª Ø§ÙØªØ¯ÙÙÙ")
    age = age.sort_values([FROZEN_VALUE, OLD_VALUE, "value"], ascending=False).rename(columns=AGING_LABELS)
    if sources.margin is None:
        audit, age = audit.drop(columns=MARGIN), age.drop(columns=MARGIN)
    if sources.cost is None:
        audit = audit.drop(columns=DECLARED_COST)
    return audit, age, abc.loc[:, list(ABC_LABELS)].rename(columns=ABC_LABELS)


def summaries(detail: pd.DataFrame, aging: pd.DataFrame, abc: pd.DataFrame, sources: Sources):
    qty, value = detail["qty"], detail["value"]
    positive, negative = detail["positive"], detail["negative"]
    age_ok = aging["age_valid"]
    old_qty = aging[AGE_DETAIL[4]]
    old_value = aging[OLD_VALUE]
    unallocated = positive & ~age_ok
    unknown_old = aging["unknown_old"]
    direct, proxy = aging["stalled_direct"], aging["stalled_proxy"]
    frozen = aging[FROZEN_VALUE]
    gross = value[positive & value.ge(0)].sum()
    quantity_totals = sources.unit is None or (
        detail.loc[positive | negative, "unit"].notna().all() and detail["unit"].nunique() <= 1
    )
    losses = detail["margin_loss"]
    margin_available = detail.loc[positive, "margin"].notna().any()
    no_margin = positive & detail["margin"].isna()
    margin_coverage = [(
        "Ø£Ø±ØµØ¯Ø© ÙÙØ¬Ø¨Ø© Ø¨ÙØ§ ÙØ§ÙØ´ ÙÙØ«ÙÙ", no_margin.sum(), None, value[no_margin & value.ge(0)].sum()
    )] if sources.margin else []
    kpis = pd.DataFrame([
        ("Ø§ÙØ³Ø¬ÙØ§Øª Ø§ÙÙÙØ±ÙØ¡Ø©", len(detail), None, None),
        ("Ø§ÙØ£ØµÙØ§Ù Ø§ÙÙØ±ÙØ¯Ø©", detail["sku"].nunique(), None, None),
        ("Ø§ÙØ£Ø±ØµØ¯Ø© Ø§ÙÙÙØ¬Ø¨Ø© ÙÙÙÙØªÙØ§ Ø§ÙØµØ§ÙØ­Ø©", positive.sum(), None, gross),
        ("Ø§ÙØ£Ø±ØµØ¯Ø© Ø§ÙØ³Ø§ÙØ¨Ø© ØªØ­Øª Ø§ÙØªØ¯ÙÙÙ", negative.sum(), qty[negative].sum(), value[negative].sum(min_count=1)),
        ("ØµØ§ÙÙ ÙÙÙ Ø§ÙÙØµØ¯Ø± Ø§ÙÙÙØ±ÙØ¡Ø©", None, None, value.sum(min_count=1)),
        ("Ø£Ø±ØµØ¯Ø© Ø£ÙØ¯Ù ÙÙ 180 ÙÙÙÙØ§", old_qty.gt(0).sum(), old_qty.sum(), old_value.sum()),
        ("Ø£Ø±ØµØ¯Ø© ÙØ¯ÙÙØ© ÙØ£Ø¯ÙØ© Ø§ÙØ­Ø±ÙØ© ØºÙØ± ÙØ§ÙÙØ©", unknown_old.sum(), old_qty[unknown_old].sum(), old_value[unknown_old].sum()),
        ("Ø±ÙÙØ¯ Ø­Ø³Ø¨ ØªØ§Ø±ÙØ® Ø§ÙØ­Ø±ÙØ©", direct.sum() if sources.movement else "ØºÙØ± ÙØªØ§Ø­",
         old_qty[direct].sum() if sources.movement else None, old_value[direct].sum() if sources.movement else None),
        ("Ø±ÙÙØ¯ Ø­Ø³Ø¨ Ø¢Ø®Ø± Ø¨ÙØ¹ ÙØªÙØ±ÙØ¯", proxy.sum() if sources.sale and sources.receipt else "ØºÙØ± ÙØªØ§Ø­",
         old_qty[proxy].sum() if sources.sale and sources.receipt else None,
         old_value[proxy].sum() if sources.sale and sources.receipt else None),
        ("Ø£Ø±ØµØ¯Ø© ÙÙØ¬Ø¨Ø© Ø¨ÙØ§ÙØ´ Ø³Ø§ÙØ¨", losses.sum() if margin_available else "ØºÙØ± ÙØªØ§Ø­",
         None, aging["loss_value"].sum() if margin_available else None),
        *margin_coverage,
        ("Ø¥Ø¬ÙØ§ÙÙ Ø§ÙØªØ¬ÙÙØ¯ Ø¨Ø¹Ø¯ Ø¥Ø²Ø§ÙØ© Ø§ÙØªØ¯Ø§Ø®Ù", frozen.gt(0).sum(), None, frozen.sum()),
        ("Ø³Ø¬ÙØ§Øª ØªØ­Øª Ø§ÙØªØ¯ÙÙÙ", detail["reasons"].ne("").sum(), None, None),
    ], columns=["Ø§ÙÙØ¤Ø´Ø±", "Ø§ÙØ¹Ø¯Ø¯", "Ø§ÙÙÙÙØ© Ø§ÙÙØ³Ø¬ÙØ©", "Ø§ÙÙÙÙØ©"])
    by_age = pd.DataFrame([
        (AGES[0], aging[AGE_DETAIL[0]].sum(), aging[AGE_DETAIL[1]].sum()),
        (AGES[1], aging[AGE_DETAIL[2]].sum(), aging[AGE_DETAIL[3]].sum()),
        (AGES[2], old_qty.sum(), old_value.sum()),
        (AGES[3], qty[unallocated].sum(), value[unallocated & value.ge(0)].sum()),
    ], columns=["Ø¹ÙØ± Ø§ÙØ±ØµÙØ¯ Ø§ÙÙÙØ¬Ø¨", "Ø§ÙÙÙÙØ© Ø§ÙÙØ³Ø¬ÙØ©", "Ø§ÙÙÙÙØ©"])
    by_age["Ø­ØµØ© Ø§ÙÙÙÙØ©"] = by_age["Ø§ÙÙÙÙØ©"].div(gross) if gross > 0 else np.nan
    if not quantity_totals:
        kpis["Ø§ÙÙÙÙØ© Ø§ÙÙØ³Ø¬ÙØ©"] = None
        by_age["Ø§ÙÙÙÙØ© Ø§ÙÙØ³Ø¬ÙØ©"] = None
    by_abc = (
        abc.groupby("value_abc").agg(Ø§ÙØ¹Ø¯Ø¯=("sku", "size"), Ø§ÙÙÙÙØ©=("value", "sum"))
        .reindex(["A", "B", "C"], fill_value=0).rename_axis("Ø§ÙÙØ¦Ø©").reset_index()
    )
    eligible_value = by_abc["Ø§ÙÙÙÙØ©"].sum()
    by_abc["Ø­ØµØ© Ø§ÙÙÙÙØ©"] = by_abc["Ø§ÙÙÙÙØ©"].div(eligible_value) if eligible_value > 0 else np.nan
    return kpis, by_age, by_abc


def excel_formats(workbook: Workbook):
    return {
        "title": workbook.add_format({
            "font_name": "Arial", "font_size": 20, "bold": True, "font_color": "#FFFFFF",
            "bg_color": "#142E45", "align": "right", "valign": "vcenter",
        }),
        "section": workbook.add_format({
            "font_name": "Arial", "font_size": 11, "bold": True, "font_color": "#FFFFFF",
            "bg_color": "#24576C", "align": "right", "valign": "vcenter",
        }),
        "header": workbook.add_format({
            "font_name": "Arial", "bold": True, "font_color": "#FFFFFF", "bg_color": "#24576C",
            "text_wrap": True, "align": "right", "valign": "vcenter",
        }),
        "note": workbook.add_format({
            "font_name": "Arial", "font_color": "#2C4C62", "bg_color": "#EAF2F6",
            "align": "right", "valign": "vcenter", "text_wrap": True,
        }),
        "number": workbook.add_format({"num_format": "#,##0.00;[Red](#,##0.00);â"}),
        "integer": workbook.add_format({"num_format": "#,##0;[Red](#,##0);â"}),
        "percent": workbook.add_format({"num_format": "0.0%"}),
        "date": workbook.add_format({"num_format": "yyyy-mm-dd"}),
    }


def excel_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.to_pydatetime() if isinstance(value, pd.Timestamp) else value
    if isinstance(value, np.generic):
        return value.item()
    return value


def write_table(worksheet: Worksheet, frame: pd.DataFrame, row: int, formats, *, table_name: str):
    if len(frame) + row + 1 > EXCEL_MAX_ROWS:
        raise ValueError(f"Ø¹Ø¯Ø¯ Ø§ÙØµÙÙÙ ÙØªØ¬Ø§ÙØ² Ø­Ø¯ Excel ÙÙ ÙØ±ÙØ© {worksheet.name}.")
    if frame.empty:
        worksheet.write(row, 0, "ÙØ§ ØªÙØ¬Ø¯ Ø³Ø¬ÙØ§Øª.", formats["note"])
        return
    columns = list(frame.columns)
    for col, heading in enumerate(columns):
        worksheet.write(row, col, str(heading), formats["header"])
    for offset, values in enumerate(frame.itertuples(index=False, name=None), start=1):
        for col, value in enumerate(values):
            value = excel_value(value)
            fmt = formats["date"] if isinstance(value, (datetime, date)) else None
            worksheet.write(row + offset, col, value, fmt)
    worksheet.add_table(row, 0, row + len(frame), len(columns) - 1, {
        "name": table_name, "style": "Table Style Medium 2",
        "columns": [{"header": str(c)} for c in columns],
        "autofilter": True, "banded_rows": True,
    })
    worksheet.freeze_panes(row + 1, 0)
    worksheet.set_column(0, len(columns) - 1, 18)
    for col, name in enumerate(columns):
        if any(token in str(name) for token in ("ÙÙÙØ©", "ÙÙÙØ©", "Ø§ÙØ±ØµÙØ¯", "ØªÙÙÙØ©", "ÙØ§ÙØ´", "Ø§ÙÙÙÙØ©")):
            worksheet.set_column(col, col, 18, formats["number"])


def write_report(path: Path, tables, summary_tables, as_of: pd.Timestamp, config: InventoryConfig):
    path.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=path.parent) as tmp:
        temp_path = Path(tmp) / path.name
        workbook = Workbook(str(temp_path), {"constant_memory": True})
        formats = excel_formats(workbook)
        try:
            overview = workbook.add_worksheet("ÙÙØ®Øµ")
            overview.right_to_left()
            overview.set_column("A:A", 38)
            overview.set_column("B:D", 22)
            overview.merge_range("A1:D1", "ØªØ­ÙÙÙ Ø£Ø¹ÙØ§Ø± Ø§ÙÙØ®Ø²ÙÙ", formats["title"])
            overview.set_row(0, 34)
            overview.write("A3", "ØªØ§Ø±ÙØ® Ø§ÙÙÙØ·Ø©", formats["section"])
            overview.write_datetime("B3", as_of.to_pydatetime(), formats["date"])
            overview.merge_range("A5:D5", "Ø§ÙÙØ¤Ø´Ø±Ø§Øª Ø§ÙØ±Ø¦ÙØ³ÙØ©", formats["section"])
            write_table(overview, summary_tables[0], 5, formats, table_name="KPIs")
            start = 8 + len(summary_tables[0])
            overview.merge_range(start, 0, start, 3, "Ø§ÙØªÙØ²ÙØ¹ Ø§ÙØ¹ÙØ±Ù", formats["section"])
            write_table(overview, summary_tables[1], start + 1, formats, table_name="AgeSummary")
            start2 = start + 5 + len(summary_tables[1])
            overview.merge_range(start2, 0, start2, 3, "ØªØµÙÙÙ ABC Ø­Ø³Ø¨ Ø§ÙÙÙÙØ©", formats["section"])
            write_table(overview, summary_tables[2], start2 + 1, formats, table_name="ABCSummary")
            for name, frame, table_name in zip(
                ("Ø§ÙØªØ¯ÙÙÙ", "Ø£Ø¹ÙØ§Ø± Ø§ÙÙØ®Ø²ÙÙ", "ØªØµÙÙÙ Ø§ÙØ£ØµÙØ§Ù"), tables,
                ("Audit", "Aging", "ABCItems"), strict=True,
            ):
                sheet = workbook.add_worksheet(name)
                sheet.right_to_left()
                sheet.merge_range(0, 0, 0, max(0, len(frame.columns) - 1), name, formats["title"])
                write_table(sheet, frame, 2, formats, table_name=table_name)
            workbook.close()
            temp_path.replace(path)
        except Exception:
            workbook.close()
            raise


def analyze_inventory(config: InventoryConfig):
    with timed("ÙØ±Ø§Ø¡Ø© Ø§ÙÙØµØ¯Ø±", config.verbose):
        as_of = snapshot_date(config)
        raw, encoding = read_inventory(config)
    with timed("ØªØ­ÙÙÙ Ø§ÙØ³Ø¬ÙØ§Øª", config.verbose):
        detail, sources = prepare_inventory(raw, config, as_of)
        aging = analyze_aging(detail, raw, config)
        abc, _, _ = classify_inventory(detail, config)
        tables = report_tables(detail, aging, abc, sources)
        summary_tables = summaries(detail, aging, abc, sources)
    with timed("ÙØªØ§Ø¨Ø© Ø§ÙØªÙØ±ÙØ±", config.verbose):
        write_report(config.output, tables, summary_tables, as_of, config)
    return {"output": config.output, "as_of": as_of.date(), "encoding": encoding,
            "records": len(raw), "audit_records": len(tables[0])}


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="ØªØ­ÙÙÙ Ø£Ø¹ÙØ§Ø± Ø§ÙÙØ®Ø²ÙÙ ÙØ¥ÙØ´Ø§Ø¡ ØªÙØ±ÙØ± Excel.")
    parser.add_argument("source", type=Path, help="ÙÙÙ CSV")
    parser.add_argument("-o", "--output", type=Path, default=Path("inventory_report.xlsx"))
    parser.add_argument("--as-of", type=date.fromisoformat, help="ØªØ§Ø±ÙØ® Ø§ÙÙÙØ·Ø© Ø¨ØµÙØºØ© YYYY-MM-DD")
    parser.add_argument("--encoding", default="auto")
    parser.add_argument("--separator", default=",")
    parser.add_argument("--date-format", default="%Y-%m-%d")
    parser.add_argument("--uom-column")
    parser.add_argument("--unit-cost-column")
    parser.add_argument("--margin-column")
    parser.add_argument("--last-movement-column")
    parser.add_argument("--margin-is-ratio", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)
    return InventoryConfig(
        source=args.source, output=args.output, as_of=args.as_of, encoding=args.encoding,
        separator=args.separator, date_format=args.date_format, uom_column=args.uom_column,
        unit_cost_column=args.unit_cost_column, margin_column=args.margin_column,
        last_movement_column=args.last_movement_column, margin_is_ratio=args.margin_is_ratio,
        verbose=not args.quiet,
    )


def main(argv=None) -> int:
    try:
        result = analyze_inventory(parse_args(argv))
    except (OSError, ValueError, pd.errors.ParserError, XlsxWriterException) as exc:
        print(f"Ø®Ø·Ø£: {exc}", file=sys.stderr)
        return 2
    print(f"Ø§ÙØªÙÙ Ø§ÙØªÙØ±ÙØ±: {result['output']} | Ø§ÙØ³Ø¬ÙØ§Øª: {result['records']} | ØªØ­Øª Ø§ÙØªØ¯ÙÙÙ: {result['audit_records']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
