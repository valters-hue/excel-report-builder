"""Cleaning rules for messy sales/order exports.

Each function takes raw values as they come out of a CSV and returns a clean
value or None, so the report can count exactly what was fixed or dropped.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional

import pandas as pd

DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
    "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
]


def parse_date(value, day_first: bool = True) -> Optional[date]:
    """Parse the date formats seen in real exports. Ambiguous 03/04/2026 is
    read day-first by default (EU); pass day_first=False for US files."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    formats = DATE_FORMATS if day_first else [f for f in DATE_FORMATS if f != "%d/%m/%Y"] + ["%d/%m/%Y"]
    if not day_first:
        formats.remove("%m/%d/%Y")
        formats.insert(3, "%m/%d/%Y")
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Excel serial number (days since 1899-12-30)
    if re.fullmatch(r"\d{5}(\.\d+)?", s):
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(s))).date()
    return None


def parse_amount(value) -> Optional[float]:
    """'€1.234,50' → 1234.5, '$1,200' → 1200.0, '(45.00)' → -45.0, '' → None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s or s.lower() in {"n/a", "na", "-", "—", "null", "none"}:
        return None
    negative = s.startswith("(") and s.endswith(")") or s.startswith("-")
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
        s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in s:
        tail = s.split(",")[-1]
        s = s.replace(",", "") if len(tail) == 3 else s.replace(",", ".")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negative else n


def clean_name(value) -> Optional[str]:
    """'  ACME   builders ltd. ' → 'Acme Builders Ltd.' (keeps short all-caps acronyms)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    words = str(value).split()
    if not words:
        return None
    out = []
    for w in words:
        if w.isupper() and len(w) <= 4 and w.isalpha() and w not in {"LTD", "INC"}:
            out.append(w)  # SIA, AS, UAB, BMW stay as they are
        else:
            out.append(w[:1].upper() + w[1:].lower())
    return " ".join(out)


def clean_category(value, aliases: dict[str, str] | None = None) -> str:
    s = re.sub(r"\s+", " ", str(value or "").strip()).lower()
    if not s or s == "nan":
        return "Uncategorised"
    aliases = aliases or {}
    return aliases.get(s, s[:1].upper() + s[1:])
