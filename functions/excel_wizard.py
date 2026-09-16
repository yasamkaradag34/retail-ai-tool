"""Deterministic, ephemeral Excel Wizard engine.

The language planner in this module only chooses a validated operation and
column roles.  Every value written to a workbook and every metric returned to
the caller is calculated locally with pandas/openpyxl.  Workbooks never touch
disk: they live in an owner-bound, size-limited, TTL cache.
"""

from __future__ import annotations

from collections import OrderedDict
from copy import copy, deepcopy
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from io import BytesIO, StringIO
import json
import math
import re
import secrets
import threading
import time
import unicodedata
from typing import Any, Callable, Iterable, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_TTL_SECONDS = 60 * 60
DEFAULT_MAX_UPLOAD_BYTES = 25 * 1024 * 1024
DEFAULT_MAX_WORKBOOKS = 100
DEFAULT_MAX_WORKBOOKS_PER_OWNER = 5
MAX_ROWS = 200_000
MAX_COLUMNS = 300
PREVIEW_ROWS = 100


CAPABILITIES = [
    {"id": "sum", "label": "TOPLA (SUM)", "group": "calculation"},
    {"id": "average", "label": "ORTALAMA (AVERAGE)", "group": "calculation"},
    {"id": "if", "label": "EĞER (IF)", "group": "formula"},
    {"id": "ifs", "label": "ÇOKEĞER (IFS)", "group": "formula"},
    {"id": "sumproduct", "label": "TOPLA.ÇARPIM (SUMPRODUCT)", "group": "calculation"},
    {"id": "sumif", "label": "ETOPLA (SUMIF)", "group": "calculation"},
    {"id": "sumifs", "label": "ÇOKETOPLA (SUMIFS)", "group": "calculation"},
    {"id": "countif", "label": "EĞERSAY (COUNTIF)", "group": "calculation"},
    {"id": "countifs", "label": "ÇOKEĞERSAY (COUNTIFS)", "group": "calculation"},
    {"id": "counta", "label": "BAĞ_DEĞ_DOLU_SAY (COUNTA)", "group": "calculation"},
    {"id": "vlookup", "label": "DÜŞEYARA (VLOOKUP)", "group": "lookup"},
    {"id": "hlookup", "label": "YATAYARA (HLOOKUP)", "group": "lookup"},
    {"id": "index_match", "label": "İNDİS + KAÇINCI (INDEX + MATCH)", "group": "lookup"},
    {"id": "lookup", "label": "ARA (LOOKUP)", "group": "lookup"},
    {"id": "xlookup", "label": "XLOOKUP", "group": "lookup"},
    {"id": "concat", "label": "BİRLEŞTİR (CONCAT)", "group": "text"},
    {"id": "textjoin", "label": "METİNBİRLEŞTİR (TEXTJOIN)", "group": "text"},
    {"id": "left", "label": "SOLDAN (LEFT)", "group": "text"},
    {"id": "right", "label": "SAĞDAN (RIGHT)", "group": "text"},
    {"id": "mid", "label": "PARÇAAL (MID)", "group": "text"},
    {"id": "upper", "label": "BÜYÜK_HARF (UPPER)", "group": "text"},
    {"id": "lower", "label": "KÜÇÜK_HARF (LOWER)", "group": "text"},
    {"id": "substitute", "label": "METNİDÜZENLE (SUBSTITUTE)", "group": "text"},
    {"id": "find", "label": "BUL (FIND)", "group": "text"},
    {"id": "len", "label": "UZUNLUK (LEN)", "group": "text"},
    {"id": "today", "label": "BUGÜN (TODAY)", "group": "date"},
    {"id": "now", "label": "ŞİMDİ (NOW)", "group": "date"},
    {"id": "date", "label": "TARİH (DATE)", "group": "date"},
    {"id": "day", "label": "GÜN (DAY)", "group": "date"},
    {"id": "month", "label": "AY (MONTH)", "group": "date"},
    {"id": "year", "label": "YIL (YEAR)", "group": "date"},
    {"id": "conditional_format", "label": "Koşullu Biçimlendirme", "group": "format"},
    {"id": "filter", "label": "Filtre", "group": "data"},
    {"id": "sort", "label": "Sırala", "group": "data"},
    {"id": "pivot", "label": "Pivot Tablo", "group": "analysis"},
    {"id": "data_validation", "label": "Veri Doğrulama", "group": "data"},
    {"id": "text_to_columns", "label": "Metni Sütunlara Dönüştür", "group": "data"},
    {"id": "chart", "label": "Grafik Oluşturma", "group": "visual"},
    {"id": "commerce_analysis", "label": "E-ticaret Performans Analizi", "group": "analysis"},
]
CAPABILITY_IDS = {item["id"] for item in CAPABILITIES}


ROLE_ALIASES = {
    "metric_column": ["revenue", "ciro", "turnover", "sales amount", "satis tutari", "gmv", "net sales", "amount", "tutar"],
    "revenue": ["revenue", "net revenue", "gross revenue", "ciro", "gelir", "hasilat", "turnover", "sales amount", "satis tutari", "gmv", "net sales", "gross sales"],
    "brand": ["brand", "brand name", "marka", "manufacturer", "vendor", "uretici"],
    "category": ["category", "category name", "kategori", "product category", "item category", "urun kategorisi", "cat1", "cat2", "taxonomy"],
    "product": ["product", "product name", "product title", "urun", "urun adi", "sku", "item", "item name", "title", "name", "model"],
    "order": ["order", "order id", "order number", "order no", "siparis", "siparis id", "siparis no", "transaction id"],
    "orders": ["orders", "order count", "order quantity", "siparis sayisi", "siparis adedi", "total orders"],
    "units": ["units", "units sold", "items sold", "quantity sold", "sales qty", "sold qty", "satilan adet", "satis adedi", "adet"],
    "price": ["price", "selling price", "sale price", "fiyat", "unit price", "birim fiyat"],
    "cost": ["cost", "unit cost", "cogs", "maliyet", "birim maliyet", "product cost"],
    "margin": ["margin", "gross margin", "profit margin", "contribution margin", "kar marji", "marj", "gross profit", "kar"],
    "stock": ["stock", "stock quantity", "stock qty", "stok", "stok adedi", "inventory", "inventory level", "on hand", "available quantity", "available qty"],
    "transactions": ["transactions", "transaction count", "purchases", "purchase count", "ecommerce purchases", "completed purchases", "sales count", "satin alma", "satin alim"],
    "views": ["pdp views", "product detail views", "product views", "detail page views", "views", "view item", "view_item", "goruntuleme", "urun goruntuleme", "item views"],
    "add_to_cart": ["add to cart", "add_to_cart", "adds to cart", "add to carts", "sepete ekleme", "sepete eklenen", "cart adds", "a2c"],
    "date": ["date", "tarih", "order date", "created at", "day"],
}


# Words that make a superficially similar heading unsafe for a role.  For
# example, ``Order Date`` must never be silently treated as an order count and
# ``Stock Quantity`` must not be selected as sold units.
ROLE_NEGATIVE_TOKENS = {
    "revenue": {"cost", "maliyet", "margin", "marj", "price", "fiyat", "qty", "quantity", "adet", "rate", "oran"},
    "metric_column": {"cost", "maliyet", "margin", "marj", "price", "fiyat", "qty", "quantity", "adet", "rate", "oran"},
    "order": {"date", "tarih", "time", "zaman", "status", "durum"},
    "orders": {"date", "tarih", "time", "zaman", "status", "durum"},
    "units": {"stock", "stok", "inventory", "available", "mevcut", "view", "goruntuleme"},
    "stock": {"sold", "satilan", "sales", "satis", "order", "siparis", "purchase", "view"},
    "transactions": {"id", "date", "tarih", "rate", "oran", "revenue", "ciro"},
    "views": {"rate", "oran", "conversion", "donusum"},
    "add_to_cart": {"rate", "oran", "conversion", "donusum"},
    "price": {"cost", "maliyet"},
    "cost": {"price", "fiyat"},
}


class ExcelWizardError(Exception):
    def __init__(self, status_code: int, code: str, message: str):
        super().__init__(message)
        self.status_code = status_code
        self.code = code
        self.message = message


@dataclass
class PendingClarification:
    clarification_id: str
    command: str
    mode: str
    operation: dict[str, Any]
    role: str
    created_at: float = field(default_factory=time.time)


@dataclass
class WorkbookSession:
    workbook_id: str
    owner: str
    filename: str
    workbook: Any
    dataframe: pd.DataFrame
    original_dataframe: pd.DataFrame
    active_sheet: str
    uploaded_bytes: int
    created_at: float = field(default_factory=time.time)
    last_access: float = field(default_factory=time.time)
    modified: bool = False
    role_mappings: dict[str, str] = field(default_factory=dict)
    operation_log: list[dict[str, Any]] = field(default_factory=list)
    pending: dict[str, PendingClarification] = field(default_factory=dict)
    lock: Any = field(default_factory=threading.RLock, repr=False, compare=False)


def _norm(value: Any) -> str:
    text = str(value or "").strip().casefold()
    text = text.translate(str.maketrans({"ı": "i", "ş": "s", "ğ": "g", "ü": "u", "ö": "o", "ç": "c"}))
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def _json_value(value: Any) -> Any:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


def dataframe_preview(df: pd.DataFrame, limit: int = PREVIEW_ROWS) -> dict[str, Any]:
    safe = df.head(limit)
    rows = [
        {str(column): _json_value(value) for column, value in row.items()}
        for row in safe.to_dict(orient="records")
    ]
    return {"columns": [str(column) for column in df.columns], "rows": rows, "total_rows": int(len(df))}


def dataframe_profile(df: pd.DataFrame, workbook: Any | None = None) -> dict[str, Any]:
    date_columns = sum(
        pd.api.types.is_datetime64_any_dtype(df[column]) or any(word in _norm(column).split() for word in ("date", "tarih"))
        for column in df.columns
    )
    numeric_columns = sum(pd.api.types.is_numeric_dtype(df[column]) for column in df.columns)
    return {
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "missing_cells": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "numeric_columns": int(numeric_columns),
        "text_columns": int(max(0, len(df.columns) - numeric_columns - date_columns)),
        "date_columns": int(date_columns),
        "sheet_names": list(workbook.sheetnames) if workbook is not None else [],
    }


def _safe_csv_cell(value: Any) -> Any:
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _workbook_from_dataframe(df: pd.DataFrame, title: str = "Data") -> Any:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = title[:31] or "Data"
    for col_index, column in enumerate(df.columns, 1):
        cell = worksheet.cell(1, col_index, str(column))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row_index, row in enumerate(df.itertuples(index=False, name=None), 2):
        for col_index, value in enumerate(row, 1):
            worksheet.cell(row_index, col_index, _json_value(value))
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    return workbook


def _select_sheet(workbook: Any) -> str:
    for worksheet in workbook.worksheets:
        if worksheet.max_row > 1 and worksheet.max_column > 0:
            return worksheet.title
    return workbook.active.title


def _dataframe_from_worksheet(worksheet: Any) -> pd.DataFrame:
    values = list(worksheet.values)
    if not values:
        return pd.DataFrame()
    raw_headers = list(values[0])
    seen: dict[str, int] = {}
    headers: list[str] = []
    for index, header in enumerate(raw_headers, 1):
        base = str(header).strip() if header not in (None, "") else f"Column_{index}"
        seen[base] = seen.get(base, 0) + 1
        headers.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    return pd.DataFrame(values[1:], columns=headers).dropna(how="all").reset_index(drop=True)


def parse_workbook(contents: bytes, filename: str) -> tuple[Any, pd.DataFrame, str]:
    suffix = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if suffix in {".xlsx", ".xlsm"}:
        try:
            workbook = load_workbook(BytesIO(contents), data_only=False, keep_vba=suffix == ".xlsm")
            sheet = _select_sheet(workbook)
            try:
                df = pd.read_excel(BytesIO(contents), sheet_name=sheet)
            except Exception:
                df = _dataframe_from_worksheet(workbook[sheet])
        except Exception as exc:
            raise ExcelWizardError(400, "invalid_workbook", "Excel dosyası açılamadı veya bozuk.") from exc
    elif suffix == ".xls":
        try:
            df = pd.read_excel(BytesIO(contents))
            workbook = _workbook_from_dataframe(df)
            sheet = workbook.active.title
        except Exception as exc:
            raise ExcelWizardError(400, "invalid_workbook", "Excel dosyası açılamadı veya bozuk.") from exc
    elif suffix == ".csv":
        decoded = None
        for encoding in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
            try:
                decoded = contents.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            raise ExcelWizardError(400, "invalid_csv", "CSV dosyasının karakter kodlaması okunamadı.")
        try:
            df = pd.read_csv(StringIO(decoded), sep=None, engine="python")
        except (pd.errors.ParserError, UnicodeDecodeError, ValueError) as exc:
            raise ExcelWizardError(400, "invalid_csv", "CSV dosyası ayrıştırılamadı.") from exc
        df = df.map(_safe_csv_cell)
        workbook = _workbook_from_dataframe(df)
        sheet = workbook.active.title
    else:
        raise ExcelWizardError(400, "unsupported_file", "Yalnızca .xlsx, .xls, .xlsm veya .csv yükleyebilirsiniz.")
    if len(df) > MAX_ROWS or len(df.columns) > MAX_COLUMNS:
        raise ExcelWizardError(413, "workbook_too_large", f"Çalışma kitabı en fazla {MAX_ROWS:,} satır ve {MAX_COLUMNS} sütun içerebilir.")
    if df.columns.empty:
        raise ExcelWizardError(400, "empty_workbook", "Çalışma kitabında başlıklı bir veri tablosu bulunamadı.")
    df.columns = [str(column) for column in df.columns]
    return workbook, df, sheet


class ExcelWorkbookStore:
    def __init__(
        self,
        ttl_seconds: int = DEFAULT_TTL_SECONDS,
        max_upload_bytes: int = DEFAULT_MAX_UPLOAD_BYTES,
        max_workbooks: int = DEFAULT_MAX_WORKBOOKS,
        max_workbooks_per_owner: int = DEFAULT_MAX_WORKBOOKS_PER_OWNER,
    ):
        self.ttl_seconds = ttl_seconds
        self.max_upload_bytes = max_upload_bytes
        self.max_workbooks = max_workbooks
        self.max_workbooks_per_owner = max_workbooks_per_owner
        self._items: OrderedDict[str, WorkbookSession] = OrderedDict()
        self._lock = threading.RLock()

    def _cleanup_locked(self) -> None:
        cutoff = time.time() - self.ttl_seconds
        expired = [key for key, record in self._items.items() if record.last_access < cutoff]
        for key in expired:
            self._items.pop(key, None)

    def cleanup(self) -> int:
        with self._lock:
            before = len(self._items)
            self._cleanup_locked()
            return before - len(self._items)

    def clear(self) -> None:
        with self._lock:
            self._items.clear()

    def expire(self, workbook_id: Optional[str] = None) -> int:
        """Mark one or all workbooks expired; useful for deterministic cleanup tests."""
        with self._lock:
            targets = [self._items[workbook_id]] if workbook_id in self._items else ([] if workbook_id else list(self._items.values()))
            for record in targets:
                record.last_access = time.time() - self.ttl_seconds - 1
            return len(targets)

    def create(self, owner: str, filename: str, contents: bytes) -> WorkbookSession:
        owner = owner.strip().casefold()
        if not owner:
            raise ExcelWizardError(401, "login_required", "Excel Wizard için oturum açın.")
        if not contents:
            raise ExcelWizardError(400, "empty_upload", "Yüklenen dosya boş.")
        if len(contents) > self.max_upload_bytes:
            raise ExcelWizardError(413, "file_too_large", f"Dosya boyutu {self.max_upload_bytes // (1024 * 1024)} MB sınırını aşıyor.")
        workbook, dataframe, active_sheet = parse_workbook(contents, filename)
        record = WorkbookSession(
            workbook_id=secrets.token_urlsafe(18),
            owner=owner,
            filename=(filename or "workbook.xlsx")[:180],
            workbook=workbook,
            dataframe=dataframe.copy(),
            original_dataframe=dataframe.copy(),
            active_sheet=active_sheet,
            uploaded_bytes=len(contents),
        )
        with self._lock:
            self._cleanup_locked()
            owner_items = [item for item in self._items.values() if item.owner == owner]
            while len(owner_items) >= self.max_workbooks_per_owner:
                oldest = min(owner_items, key=lambda item: item.last_access)
                self._items.pop(oldest.workbook_id, None)
                owner_items.remove(oldest)
            while len(self._items) >= self.max_workbooks:
                self._items.popitem(last=False)
            self._items[record.workbook_id] = record
        return record

    def get(self, owner: str, workbook_id: str) -> WorkbookSession:
        with self._lock:
            self._cleanup_locked()
            record = self._items.get(workbook_id)
            if record is None or record.owner != owner.strip().casefold():
                raise ExcelWizardError(404, "workbook_not_found", "Çalışma kitabı bulunamadı veya oturum süresi doldu.")
            record.last_access = time.time()
            self._items.move_to_end(workbook_id)
            return record

    def delete(self, owner: str, workbook_id: str) -> bool:
        with self._lock:
            record = self._items.get(workbook_id)
            if record is None or record.owner != owner.strip().casefold():
                raise ExcelWizardError(404, "workbook_not_found", "Çalışma kitabı bulunamadı veya oturum süresi doldu.")
            self._items.pop(workbook_id, None)
            return True

    def count(self) -> int:
        with self._lock:
            self._cleanup_locked()
            return len(self._items)


excel_workbook_store = ExcelWorkbookStore()


def state_payload(record: WorkbookSession) -> dict[str, Any]:
    profile = dataframe_profile(record.dataframe, record.workbook)
    return {
        "status": "success",
        "workbook_id": record.workbook_id,
        "filename": record.filename,
        "total_rows": profile["rows"],
        "total_columns": profile["columns"],
        "profile": profile,
        "preview": dataframe_preview(record.dataframe),
        "is_modified": record.modified,
        "role_mappings": dict(record.role_mappings),
        "capabilities": CAPABILITIES,
        "expires_in_seconds": max(0, int(record.last_access + excel_workbook_store.ttl_seconds - time.time())),
    }


def _is_numeric_column(df: pd.DataFrame, column: str) -> bool:
    series = df[column]
    if pd.api.types.is_numeric_dtype(series):
        return True
    populated = series.dropna()
    if populated.empty:
        return False
    # CSV exports often carry numeric measures as text.  Treat a column as a
    # measure only when almost every populated cell converts cleanly.
    return float(pd.to_numeric(populated, errors="coerce").notna().mean()) >= 0.9


def _column_scores(df: pd.DataFrame, phrase: str, aliases: Iterable[str] = (), role: str = "") -> list[tuple[float, str]]:
    target = _norm(phrase)
    target_words = set(target.split())
    alias_values = {_norm(alias) for alias in aliases if alias}
    scored: list[tuple[float, str]] = []
    for column in df.columns:
        normalized = _norm(column)
        words = set(normalized.split())
        negatives = ROLE_NEGATIVE_TOKENS.get(role, set())
        if negatives & words or any(token and token in normalized for token in negatives):
            continue
        score = 0.0
        if normalized == target and target:
            score = 100.0
        elif target and (target in normalized or normalized in target):
            score = 82.0
        elif target_words and words:
            score = 65.0 * len(target_words & words) / max(len(target_words), len(words))
        if normalized in alias_values:
            score = max(score, 96.0)
        elif any(alias and (alias in normalized or normalized in alias) for alias in alias_values):
            score = max(score, 74.0)
        if score:
            scored.append((score, str(column)))
    return sorted(scored, key=lambda item: (-item[0], item[1]))


def resolve_column(record: WorkbookSession, role: str, requested: Any = None, numeric: bool = False) -> tuple[Optional[str], list[str]]:
    df = record.dataframe
    learned = record.role_mappings.get(role)
    if learned in df.columns and requested in (None, ""):
        return learned, []
    phrase = str(requested or role)
    if requested not in (None, ""):
        exact = [str(column) for column in df.columns if _norm(column) == _norm(requested)]
        if len(exact) == 1 and (not numeric or _is_numeric_column(df, exact[0])):
            return exact[0], []
    aliases = ROLE_ALIASES.get(role, []) + ([str(requested)] if requested else [])
    candidates = _column_scores(df, phrase, aliases, role=role)
    if numeric:
        numeric_names = {str(column) for column in df.columns if _is_numeric_column(df, str(column))}
        candidates = [item for item in candidates if item[1] in numeric_names]
        if not candidates:
            candidates = [(50.0, column) for column in sorted(numeric_names)]
    if not candidates:
        return None, []
    best_score = candidates[0][0]
    # A canonical heading such as Product, Category, Units or Stock is more
    # specific than a related alternative such as SKU or Quantity.  Revenue is
    # intentionally stricter because Revenue vs Net Revenue changes the money
    # total materially and must be confirmed by the user.
    if (
        requested in (None, "")
        and role not in {"revenue", "metric_column"}
        and best_score >= 96.0
        and (len(candidates) == 1 or candidates[1][0] < best_score)
    ):
        return candidates[0][1], []
    if requested in (None, "") and role in ROLE_ALIASES:
        semantic = [column for score, column in candidates if score >= 70.0]
        if len(semantic) > 1:
            return None, semantic[:8]
    close = [column for score, column in candidates if score >= max(50.0, best_score - 8.0)]
    if len(close) == 1:
        return close[0], []
    return None, close[:8]


def _clarification(record: WorkbookSession, command: str, mode: str, operation: dict[str, Any], role: str, candidates: list[str], question: str) -> dict[str, Any]:
    clarification_id = secrets.token_urlsafe(12)
    record.pending[clarification_id] = PendingClarification(clarification_id, command, mode, operation, role)
    return {
        "status": "needs_clarification",
        "question": question,
        "role": role,
        "options": [
            {
                "label": column,
                "value": column,
                "description": f"{column} sütununu {role.replace('_', ' ')} olarak kullan",
            }
            for column in candidates
        ],
        "clarification_id": clarification_id,
    }


def _extract_explicit_column(command: str, df: pd.DataFrame) -> Optional[str]:
    normalized = _norm(command)
    exact = [str(column) for column in df.columns if _norm(column) and _norm(column) in normalized]
    return max(exact, key=len) if exact else None


def _mentioned_columns(command: str, df: pd.DataFrame) -> list[str]:
    question = _norm(command)
    compact_question = question.replace(" ", "")
    found: list[tuple[int, str]] = []
    for column in df.columns:
        normalized = _norm(column)
        if not normalized:
            continue
        position = question.find(normalized)
        if position < 0 and normalized.replace(" ", "") in compact_question:
            position = compact_question.find(normalized.replace(" ", ""))
        if position >= 0:
            found.append((position, str(column)))
    return [column for _, column in sorted(found, key=lambda item: (item[0], -len(item[1])))]


def _value_mentioned(series: pd.Series, command: str) -> Optional[Any]:
    question = _norm(command)
    values = series.dropna().unique().tolist()
    values.sort(key=lambda item: len(str(item)), reverse=True)
    for value in values[:1000]:
        normalized = _norm(value)
        if len(normalized) > 1 and re.search(r"(?:^|\s)" + re.escape(normalized) + r"(?:$|\s)", question):
            return value
    return None


def _numeric_criterion(column: str, command: str) -> Optional[dict[str, Any]]:
    question = _norm(command)
    column_name = _norm(column)
    position = question.find(column_name)
    segment = question[position:position + 100] if position >= 0 else question
    number = re.search(r"(-?\d+(?:[.,]\d+)?)", segment)
    if not number:
        return None
    value = float(number.group(1).replace(",", "."))
    if value.is_integer():
        value = int(value)
    tail = segment[number.end():number.end() + 45]
    if re.search(r"en\s+az|ve\s+uzeri|ve\s+ustu|greater\s+than\s+or\s+equal", tail):
        operator = ">="
    elif re.search(r"en\s+fazla|ve\s+alti|less\s+than\s+or\s+equal", tail):
        operator = "<="
    elif re.search(r"kucuk|az|alt|less|below", tail):
        operator = "<"
    elif re.search(r"buyuk|fazla|ust|greater|above", tail):
        operator = ">"
    elif re.search(r"esit|equal", tail):
        operator = "="
    else:
        return None
    return {"column": column, "operator": operator, "value": value}


def _criteria_from_command(record: WorkbookSession, command: str, exclude: Iterable[str] = ()) -> list[dict[str, Any]]:
    excluded = set(exclude)
    mentioned = _mentioned_columns(command, record.dataframe)
    criteria: list[dict[str, Any]] = []
    for column in mentioned:
        if column in excluded:
            continue
        if pd.api.types.is_numeric_dtype(record.dataframe[column]):
            criterion = _numeric_criterion(column, command)
            if criterion:
                criteria.append(criterion)
        else:
            value = _value_mentioned(record.dataframe[column], command)
            if value is not None:
                criteria.append({"column": column, "operator": "=", "value": _json_value(value)})
    return criteria


def _target_name(source: str, suffix: str) -> str:
    return re.sub(r"\s+", "_", f"{source}_{suffix}")[:120]


HINT_TO_OPERATION = {
    "index": "index_match", "match": "index_match",
    "conditional-formatting": "conditional_format",
    "validation": "data_validation", "text-to-columns": "text_to_columns",
}


def _enrich_operation(operation: dict[str, Any], command: str, record: WorkbookSession, hint: str = "") -> dict[str, Any]:
    operation_id = operation["id"]
    mentioned = _mentioned_columns(command, record.dataframe)
    numeric = [column for column in mentioned if pd.api.types.is_numeric_dtype(record.dataframe[column])]
    text = [column for column in mentioned if not pd.api.types.is_numeric_dtype(record.dataframe[column])]
    q = _norm(command)

    if operation_id in {"sum", "average", "counta"} and mentioned:
        operation["column"] = numeric[0] if operation_id != "counta" and numeric else mentioned[0]
    elif operation_id == "sumproduct":
        operation["columns"] = numeric
    elif operation_id in {"sumif", "sumifs", "countif", "countifs"}:
        sum_column = numeric[-1] if operation_id.startswith("sum") and numeric else None
        if sum_column:
            operation["sum_column"] = sum_column
        operation["criteria"] = _criteria_from_command(record, command, exclude=[sum_column] if sum_column else [])
    elif operation_id == "if" and mentioned:
        source = mentioned[0]
        operation["source_column"] = source
        operation["target_column"] = _target_name(source, "IF")
        criterion = _numeric_criterion(source, command)
        if criterion:
            operation.update({"operator": criterion["operator"], "value": criterion["value"]})
        raw = command.casefold()
        result_match = re.search(r"(?:küçükse|kucukse|büyükse|buyukse|eşitse|esitse)\s+([^,\s]+).*?(?:değilse|degilse|otherwise)\s+([^,\s]+)", raw)
        if result_match:
            operation["true_value"], operation["false_value"] = result_match.group(1).title(), result_match.group(2).title()
    elif operation_id == "ifs" and mentioned:
        source = numeric[0] if numeric else mentioned[0]
        values = pd.to_numeric(record.dataframe[source], errors="coerce").dropna()
        if not values.empty:
            low, high = float(values.quantile(1 / 3)), float(values.quantile(2 / 3))
            operation.update({
                "source_column": source,
                "target_column": _target_name(source, "Segment"),
                "cases": [
                    {"operator": ">", "value": high, "result": "Yüksek"},
                    {"operator": ">=", "value": low, "result": "Orta"},
                ],
                "default": "Düşük",
            })
    elif operation_id in {"concat", "textjoin"}:
        operation["columns"] = mentioned
        operation["delimiter"] = "-" if "tire" in q or "hyphen" in q else (", " if "virgul" in q or "comma" in q else " ")
        operation["target_column"] = "_".join(mentioned[:2]) + "_Combined" if mentioned else "Birleştirilmiş_Metin"
    elif operation_id in {"left", "right", "mid", "upper", "lower", "substitute", "find", "len", "day", "month", "year"} and mentioned:
        source = mentioned[0]
        operation["source_column"] = source
        operation["target_column"] = _target_name(source, operation_id.upper())
        numbers = [int(value) for value in re.findall(r"\b\d+\b", q)]
        if operation_id in {"left", "right"} and numbers:
            operation["count"] = numbers[0] if operation_id == "left" else numbers[-1]
        elif operation_id == "mid" and numbers:
            operation["start"], operation["count"] = (numbers[-2], numbers[-1]) if len(numbers) > 1 else (1, numbers[-1])
        elif operation_id == "substitute":
            change = re.search(r"\b([^\s]+)\s+(?:ifadesini|metnini|text)\s+([^\s]+)\s+(?:ile|with)", command, flags=re.IGNORECASE)
            if change:
                operation["old"], operation["new"] = change.group(1), change.group(2)
        elif operation_id == "find":
            quoted = re.findall(r"['\"]([^'\"]+)['\"]", command)
            operation["text"] = quoted[0] if quoted else "-"
    elif operation_id == "date":
        for role, aliases in (("year_column", ["year", "yil"]), ("month_column", ["month", "ay"]), ("day_column", ["day", "gun"])):
            resolved, _ = resolve_column(record, role, aliases[0])
            if resolved:
                operation[role] = resolved
        operation["target_column"] = "Tarih"
    elif operation_id == "filter":
        operation["criteria"] = _criteria_from_command(record, command)
    elif operation_id == "sort":
        sort_column = numeric[-1] if numeric else (mentioned[-1] if mentioned else None)
        if sort_column:
            operation["column"] = sort_column
        operation["ascending"] = any(word in q for word in ["kucukten buyuge", "artan", "ascending"])
    elif operation_id == "conditional_format" and mentioned:
        operation["column"] = numeric[0] if numeric else mentioned[0]
        if "ortalama" in q or "average" in q:
            operation["threshold"] = "mean"
    elif operation_id == "data_validation" and mentioned:
        operation["column"] = mentioned[0]
        value_match = re.search(r"(?:için|icin|for)\s+(.+?)\s+(?:seçenekli|secenekli|options?|değerli|degerli)", command, flags=re.IGNORECASE)
        if value_match:
            operation["values"] = [part.strip(" .") for part in re.split(r"\s+(?:ve|and)\s+|,", value_match.group(1)) if part.strip(" .")]
    elif operation_id == "text_to_columns" and mentioned:
        operation["column"] = mentioned[0]
        operation["delimiter"] = "-" if "tire" in q or "hyphen" in q else (";" if "noktali virgul" in q else (" " if "bosluk" in q else ","))
    elif operation_id == "pivot":
        if text:
            operation["index"] = text[0]
        if len(text) > 1:
            operation["columns"] = text[1]
        if numeric:
            operation["values"] = numeric[-1]
        operation["aggfunc"] = "mean" if "ortalama" in q else ("count" if "say" in q else "sum")
    elif operation_id == "chart":
        operation["category_column"] = text[0] if text else (mentioned[0] if mentioned else None)
        operation["value_columns"] = numeric
        operation["chart_type"] = "pie" if "pasta" in q or "pie" in q else ("line" if "cizgi" in q or "line" in q else "bar")
    elif operation_id == "commerce_analysis":
        if any(word in q for word in ["funnel", "pdp", "add to cart", "sepete"]):
            operation["analysis_type"] = "funnel"
            operation["group_role"] = "product"
        elif any(word in q for word in ["stok", "stock", "inventory"]):
            operation["analysis_type"] = "stock"
            operation["group_role"] = "product"
        elif any(word in q for word in ["kategori", "category"]):
            operation["analysis_type"] = "category"
            operation["group_role"] = "category"
        elif any(word in q for word in ["marka", "brand"]):
            operation["analysis_type"] = "brand"
            operation["group_role"] = "brand"
        else:
            operation["analysis_type"] = "ranking"
        if numeric and operation["analysis_type"] not in {"funnel"}:
            operation["metric_column"] = numeric[-1]
        if text:
            preferred_role = operation.get("group_role")
            resolved, _ = resolve_column(record, str(preferred_role or "product"), requested=None)
            operation["group_column"] = resolved if resolved in text else text[0]
    elif operation_id in {"vlookup", "hlookup", "index_match", "lookup", "xlookup"}:
        if mentioned:
            operation["source_column"] = mentioned[0]
        if operation_id == "index_match" and (hint == "match" or "konum" in q or "kacinci" in q):
            source = operation.get("source_column")
            if source:
                operation["match_value"] = _value_mentioned(record.dataframe[source], command)
    return operation


def build_operation_from_command(command: str, mode: str, record: WorkbookSession, operation_hint: str = "") -> dict[str, Any]:
    q = _norm(command)
    hint = str(operation_hint or "").strip().lower()
    mapped_hint = HINT_TO_OPERATION.get(hint, hint.replace("-", "_"))
    # Longer/more specific phrases come first.
    keyword_map = [
        ("text_to_columns", ["metni sutunlara", "text to columns"]),
        ("conditional_format", ["kosullu bicimlendirme", "conditional formatting"]),
        ("data_validation", ["veri dogrulama", "data validation", "dropdown", "acilir liste"]),
        ("sumproduct", ["topla carpim", "sumproduct"]),
        ("sumifs", ["coketopla", "sumifs"]),
        ("countifs", ["cokegersay", "countifs"]),
        ("sumif", ["etopla", "sumif"]),
        ("countif", ["egersay", "countif"]),
        ("counta", ["bag deg dolu say", "counta", "dolu hucre"]),
        ("xlookup", ["xlookup", "yeniden ara"]),
        ("vlookup", ["duseyara", "vlookup"]),
        ("hlookup", ["yatayara", "hlookup"]),
        ("index_match", ["indis", "index match", "kacinci"]),
        ("textjoin", ["metinbirlestir", "textjoin"]),
        ("substitute", ["metniduzenle", "substitute", "degistir"]),
        ("concat", ["birlestir", "concat"]),
        ("average", ["ortalama", "average", "mean"]),
        ("pivot", ["pivot"]),
        ("chart", ["grafik", "chart"]),
        ("sort", ["sirala", "buyukten kucuge", "kucukten buyuge", "ascending", "descending"]),
        ("filter", ["filtre", "filter"]),
        ("ifs", ["cokeger", "ifs"]),
        ("if", ["eger", " if "]),
        ("upper", ["buyuk harf", "upper"]),
        ("lower", ["kucuk harf", "lower"]),
        ("left", ["soldan", "left"]),
        ("right", ["sagdan", "right"]),
        ("mid", ["parcaal", "mid"]),
        ("len", ["uzunluk", "len"]),
        ("find", ["bul", "find"]),
        ("today", ["bugun", "today"]),
        ("now", ["simdi", "now"]),
        ("date", ["tarih kolonu olustur", "date formula", "create date"]),
        ("day", ["gununu", "gun kolonu", " day"]),
        ("month", ["ayini", "ay kolonu", " month"]),
        ("year", ["yilini", "yil kolonu", " year"]),
        ("sum", ["toplam", "topla", "sum"]),
    ]
    operation_id = mapped_hint if mapped_hint in CAPABILITY_IDS else next((op for op, words in keyword_map if any(word in q for word in words)), None)
    commerce_words = ["en performansli", "en iyi satan", "en cok satan", "marka bazinda", "kategori bazinda", "performans", "revenue", "ciro", "satis"]
    if mode == "commerce" or (operation_id is None and any(word in q for word in commerce_words)):
        operation_id = "commerce_analysis"
    if operation_id is None:
        operation_id = "commerce_analysis" if mode == "commerce" else "sum"
    operation: dict[str, Any] = {"id": operation_id}
    explicit = _extract_explicit_column(command, record.dataframe)
    if explicit:
        operation["column"] = explicit
    if operation_id == "sort":
        operation["ascending"] = any(word in q for word in ["kucukten buyuge", "artan", "ascending"])
    if operation_id == "chart":
        operation["chart_type"] = "pie" if "pie" in q or "pasta" in q else ("line" if "line" in q or "cizgi" in q else "bar")
    if operation_id == "commerce_analysis":
        operation["top_n"] = 10
        top_match = re.search(r"(?:ilk|top)\s*(\d+)", q)
        if top_match:
            operation["top_n"] = min(100, max(1, int(top_match.group(1))))
        if any(word in q for word in ["funnel", "pdp", "add to cart", "sepete"]):
            operation["analysis_type"] = "funnel"
            operation["group_role"] = "product"
        elif any(word in q for word in ["stok", "stock", "inventory"]):
            operation["analysis_type"] = "stock"
            operation["group_role"] = "product"
        elif any(word in q for word in ["marka", "brand"]):
            operation["analysis_type"] = "brand"
            operation["group_role"] = "brand"
        elif any(word in q for word in ["kategori", "category"]):
            operation["analysis_type"] = "category"
            operation["group_role"] = "category"
        elif any(word in q for word in ["urun", "product", "sku"]):
            operation["analysis_type"] = "ranking"
            operation["group_role"] = "product"
    operation = _enrich_operation(operation, command, record, hint)

    # Gallery groups intentionally request several deterministic changes.
    if hint == "text-slice" or all(word in q for word in ("soldan", "sagdan", "ortadaki")):
        source = (_mentioned_columns(command, record.dataframe) or [operation.get("column")])[0]
        numbers = [int(value) for value in re.findall(r"\b\d+\b", q)]
        left_count, right_count, mid_count = (numbers + [3, 4, 5])[:3]
        operation = {"_batch": [
            {"id": "left", "source_column": source, "target_column": _target_name(source, "Left"), "count": left_count},
            {"id": "right", "source_column": source, "target_column": _target_name(source, "Right"), "count": right_count},
            {"id": "mid", "source_column": source, "target_column": _target_name(source, "Mid"), "start": 2, "count": mid_count},
        ]}
    elif hint == "case" or ("buyuk harf" in q and "kucuk harf" in q):
        columns = _mentioned_columns(command, record.dataframe)
        if columns:
            second = columns[1] if len(columns) > 1 else columns[0]
            operation = {"_batch": [
                {"id": "upper", "source_column": columns[0], "target_column": _target_name(columns[0], "Upper")},
                {"id": "lower", "source_column": second, "target_column": _target_name(second, "Lower")},
            ]}
    elif hint == "substitute-find" or ("degistir" in q and "konum" in q):
        columns = _mentioned_columns(command, record.dataframe)
        if columns:
            operation = {"_batch": [
                _enrich_operation({"id": "substitute"}, command, record, hint),
                _enrich_operation({"id": "find"}, command, record, hint),
            ]}
    elif hint == "today-now" or ("bugun" in q and "simdi" in q):
        operation = {"_batch": [{"id": "today", "target_column": "Bugün"}, {"id": "now", "target_column": "Şimdi"}]}
    elif hint == "date-parts" or all(word in q for word in ("gun", "ay", "yil")) and "ayri" in q:
        source, _ = resolve_column(record, "date", requested=(_mentioned_columns(command, record.dataframe) or [None])[0])
        if source:
            operation = {"_batch": [
                {"id": "day", "source_column": source, "target_column": _target_name(source, "Day")},
                {"id": "month", "source_column": source, "target_column": _target_name(source, "Month")},
                {"id": "year", "source_column": source, "target_column": _target_name(source, "Year")},
            ]}
    elif hint == "filter-sort" or ("filtre" in q and "sirala" in q):
        filter_plan = _enrich_operation({"id": "filter"}, command, record, hint)
        sort_plan = _enrich_operation({"id": "sort"}, command, record, hint)
        operation = {"_batch": [filter_plan, sort_plan]}
    return operation


def _infer_operation(command: str, mode: str, record: WorkbookSession) -> dict[str, Any]:
    return build_operation_from_command(command, mode, record)


def _strict_ai_plan(planner: Optional[Callable[[str, list[dict[str, Any]], list[str]], Any]], command: str, record: WorkbookSession) -> Optional[dict[str, Any]]:
    if planner is None:
        return None
    try:
        raw = planner(command, CAPABILITIES, [str(column) for column in record.dataframe.columns])
        if isinstance(raw, str):
            match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
            raw = json.loads(match.group(0) if match else raw)
        if not isinstance(raw, dict):
            return None
        operation_id = str(raw.get("id") or raw.get("type") or "").strip().lower()
        if operation_id not in CAPABILITY_IDS:
            return None
        allowed_columns = {str(column) for column in record.dataframe.columns}
        for key, value in raw.items():
            if key.endswith("_column") or key == "column":
                if value is not None and str(value) not in allowed_columns:
                    return None
            if key in {"columns", "value_columns", "source_columns"}:
                values = [value] if key == "columns" and operation_id == "pivot" and isinstance(value, str) else value
                if not isinstance(values, list) or any(str(item) not in allowed_columns for item in values):
                    return None
        raw["id"] = operation_id
        return raw
    except Exception:
        return None


def _criteria_mask(df: pd.DataFrame, criteria: list[dict[str, Any]]) -> pd.Series:
    mask = pd.Series(True, index=df.index)
    for criterion in criteria:
        column = str(criterion.get("column") or "")
        if column not in df.columns:
            raise ExcelWizardError(400, "invalid_column", f"'{column}' sütunu bulunamadı.")
        operator = str(criterion.get("operator") or "=").strip().lower()
        value = criterion.get("value")
        series = df[column]
        if operator in {"=", "==", "eq", "equals", "equal"}:
            current = series.astype(str).str.casefold() == str(value).casefold()
        elif operator in {"!=", "<>", "ne"}:
            current = series.astype(str).str.casefold() != str(value).casefold()
        elif operator in {">", ">=", "<", "<="}:
            left = pd.to_numeric(series, errors="coerce")
            right = float(value)
            current = {">": left > right, ">=": left >= right, "<": left < right, "<=": left <= right}[operator]
        elif operator in {"contains", "icerir"}:
            current = series.astype(str).str.contains(str(value), case=False, na=False, regex=False)
        elif operator in {"starts_with", "baslar"}:
            current = series.astype(str).str.casefold().str.startswith(str(value).casefold(), na=False)
        elif operator in {"ends_with", "biter"}:
            current = series.astype(str).str.casefold().str.endswith(str(value).casefold(), na=False)
        else:
            raise ExcelWizardError(400, "invalid_operator", f"Desteklenmeyen koşul operatörü: {operator}")
        mask &= current.fillna(False)
    return mask


def _excel_literal(value: Any) -> str:
    if value is None:
        return '""'
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    return '"' + str(value).replace('"', '""') + '"'


def _operator_formula(ref: str, operator: str, value: Any) -> str:
    operator = {"eq": "=", "ne": "<>", "==": "=", "!=": "<>"}.get(operator, operator)
    if operator in {"contains", "icerir"}:
        return f'ISNUMBER(SEARCH({_excel_literal(value)},{ref}))'
    return f"{ref}{operator}{_excel_literal(value)}"


def _ensure_target_column(record: WorkbookSession, target: str) -> tuple[int, str]:
    target = str(target or "").strip()
    if not target:
        raise ExcelWizardError(400, "target_column_required", "Oluşturulacak sütunun adını belirtin.")
    worksheet = record.workbook[record.active_sheet]
    if target in record.dataframe.columns:
        index = list(record.dataframe.columns).index(target) + 1
    else:
        index = len(record.dataframe.columns) + 1
        record.dataframe[target] = None
        worksheet.cell(1, index, target)
        if index > 1:
            worksheet.cell(1, index)._style = copy(worksheet.cell(1, index - 1)._style)
    return index, get_column_letter(index)


def _write_formula_column(record: WorkbookSession, target: str, formulas: list[str], values: Iterable[Any]) -> None:
    column_index, _ = _ensure_target_column(record, target)
    worksheet = record.workbook[record.active_sheet]
    values_list = list(values)
    record.dataframe[target] = values_list
    for row_offset, formula in enumerate(formulas, 2):
        cell = worksheet.cell(row_offset, column_index, formula)
        if column_index > 1:
            cell._style = copy(worksheet.cell(row_offset, column_index - 1)._style)
    worksheet.auto_filter.ref = f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
    record.workbook.calculation.fullCalcOnLoad = True
    record.workbook.calculation.forceFullCalc = True


def _column_ref(record: WorkbookSession, column: str, row: int) -> str:
    if column not in record.dataframe.columns:
        raise ExcelWizardError(400, "invalid_column", f"'{column}' sütunu bulunamadı.")
    return f"{get_column_letter(list(record.dataframe.columns).index(column) + 1)}{row}"


def _numeric_series(record: WorkbookSession, column: str) -> pd.Series:
    if column not in record.dataframe.columns:
        raise ExcelWizardError(400, "invalid_column", f"'{column}' sütunu bulunamadı.")
    values = pd.to_numeric(record.dataframe[column], errors="coerce")
    if not values.notna().any():
        raise ExcelWizardError(400, "numeric_column_required", f"'{column}' sütunu sayısal değer içermiyor.")
    return values


def _resolve_required(
    record: WorkbookSession,
    command: str,
    mode: str,
    operation: dict[str, Any],
    role: str,
    param: str,
    numeric: bool = False,
) -> tuple[Optional[str], Optional[dict[str, Any]]]:
    # A generic ``column`` belongs only to operations whose actual parameter is
    # named ``column``.  Reusing it for group/views/stock roles can accidentally
    # select Revenue for every role when a prompt mentions Revenue first.
    requested = operation.get(param)
    if requested in (None, ""):
        # group_column is shared by Brand, Category and Product analyses; reuse
        # the learned semantic role instead of leaking the previous analysis'
        # group into the next one.
        requested = record.role_mappings.get(role) if param == "group_column" else record.role_mappings.get(param)
    if param == "column" and requested in (None, ""):
        requested = operation.get("column")
    resolved, candidates = resolve_column(record, role, requested=requested, numeric=numeric)
    if resolved:
        operation[param] = resolved
        # These mappings live only inside the owner-bound WorkbookSession and
        # disappear with the temporary workbook.
        record.role_mappings[role] = resolved
        record.role_mappings[param] = resolved
        return resolved, None
    if not candidates:
        candidates = [str(column) for column in record.dataframe.columns if not numeric or pd.api.types.is_numeric_dtype(record.dataframe[column])][:8]
    return None, _clarification(
        record,
        command,
        mode,
        operation,
        param,
        candidates,
        f"{role.replace('_', ' ').title()} için hangi sütunu kullanmalıyım?",
    )


def _sync_values(record: WorkbookSession) -> None:
    """Write a reshaped dataframe back without replacing the workbook/sheet."""
    worksheet = record.workbook[record.active_sheet]
    old_max_row, old_max_col = worksheet.max_row, worksheet.max_column
    for col_index, column in enumerate(record.dataframe.columns, 1):
        worksheet.cell(1, col_index, str(column))
    for row_index, row in enumerate(record.dataframe.itertuples(index=False, name=None), 2):
        for col_index, value in enumerate(row, 1):
            worksheet.cell(row_index, col_index, _json_value(value))
    for row in worksheet.iter_rows(min_row=len(record.dataframe) + 2, max_row=old_max_row, min_col=1, max_col=old_max_col):
        for cell in row:
            cell.value = None
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(record.dataframe.columns))}{len(record.dataframe) + 1}"


def _analysis_range(record: WorkbookSession, column: str) -> str:
    letter = get_column_letter(list(record.dataframe.columns).index(column) + 1)
    sheet = record.active_sheet.replace("'", "''")
    return f"'{sheet}'!${letter}$2:${letter}${len(record.dataframe) + 1}"


def _criterion_formula_value(criterion: dict[str, Any]) -> str:
    operator = str(criterion.get("operator") or "=")
    operator = {"eq": "=", "equals": "=", "equal": "=", "ne": "<>", "!=": "<>"}.get(operator, operator)
    value = criterion.get("value")
    if operator in {"contains", "icerir"}:
        return _excel_literal(f"*{value}*")
    if operator in {"starts_with", "baslar"}:
        return _excel_literal(f"{value}*")
    if operator in {"ends_with", "biter"}:
        return _excel_literal(f"*{value}")
    return _excel_literal(f"{operator}{value}" if operator not in {"=", "=="} else value)


def _write_analysis_result(record: WorkbookSession, operation_id: str, columns: list[str], value: Any, formula: str, command: str) -> None:
    # Keep deterministic aggregate results in one visible audit sheet.  The
    # calculated value makes the download immediately useful, while the Excel
    # formula lets Excel recalculate it when source rows change later.
    sheet_name = "DataProvido_Results"
    if sheet_name not in record.workbook.sheetnames:
        worksheet = record.workbook.create_sheet(sheet_name)
        headers = ["Timestamp", "Operation", "Columns", "Calculated Result", "Excel Formula", "Command"]
        for index, header in enumerate(headers, 1):
            cell = worksheet.cell(1, index, header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        worksheet.freeze_panes = "A2"
    worksheet = record.workbook[sheet_name]
    row = worksheet.max_row + 1
    worksheet.cell(row, 1, datetime.now(timezone.utc).replace(microsecond=0).isoformat())
    worksheet.cell(row, 2, operation_id.upper())
    worksheet.cell(row, 3, ", ".join(columns))
    worksheet.cell(row, 4, _json_value(value))
    worksheet.cell(row, 5, formula)
    worksheet.cell(row, 6, command[:500])
    for column, width in {"A": 27, "B": 18, "C": 32, "D": 20, "E": 65, "F": 65}.items():
        worksheet.column_dimensions[column].width = width
    record.workbook.calculation.fullCalcOnLoad = True
    record.workbook.calculation.forceFullCalc = True


def _execute_aggregate(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    operation_id = operation["id"]

    def finish(summary: str, value: Any, columns: list[str], formula: str) -> dict[str, Any]:
        mutation = mode == "operations"
        if mutation:
            _write_analysis_result(record, operation_id, columns, value, formula, command)
        return {"summary": summary, "metrics": {operation_id: value}, "columns_used": columns, "is_mutation": mutation}

    if operation_id == "sumproduct":
        columns = operation.get("columns") or []
        if isinstance(columns, str):
            columns = [columns]
        if operation.get("sumproduct_column"):
            columns = list(dict.fromkeys([*columns, str(operation["sumproduct_column"])]))
            operation["columns"] = columns
            operation.pop("sumproduct_column", None)
        if len(columns) < 2:
            numeric = [str(column) for column in record.dataframe.columns if pd.api.types.is_numeric_dtype(record.dataframe[column]) and str(column) not in columns]
            known = f" Seçili: {', '.join(columns)}." if columns else ""
            return _clarification(record, command, mode, operation, "sumproduct_column", numeric[:8], f"TOPLA.ÇARPIM için {2 - len(columns)} sayısal sütun daha seçin.{known}")
        series = [_numeric_series(record, str(column)).fillna(0) for column in columns]
        result = series[0]
        for current in series[1:]:
            result = result * current
        value = float(result.sum())
        formula = "=SUMPRODUCT(" + ",".join(_analysis_range(record, column) for column in columns) + ")"
        return finish(f"{', '.join(columns)} sütunlarının TOPLA.ÇARPIM sonucu {value:,.2f}.", value, columns, formula)

    if operation_id in {"sumif", "sumifs", "countif", "countifs"}:
        criteria = operation.get("criteria") or []
        if not criteria:
            raise ExcelWizardError(400, "criteria_required", "Koşullu hesaplama için en az bir koşul belirtin.")
        mask = _criteria_mask(record.dataframe, criteria)
        if operation_id.startswith("count"):
            value = int(mask.sum())
            columns = [str(item["column"]) for item in criteria]
            function = "COUNTIF" if operation_id == "countif" else "COUNTIFS"
            arguments = []
            for item in criteria:
                arguments.extend([_analysis_range(record, str(item["column"])), _criterion_formula_value(item)])
            formula = f"={function}({','.join(arguments)})"
        else:
            sum_column, clarification = _resolve_required(record, command, mode, operation, "metric_column", "sum_column", numeric=True)
            if clarification:
                return clarification
            value = float(_numeric_series(record, sum_column)[mask].sum())
            columns = [sum_column] + [str(item["column"]) for item in criteria]
            if operation_id == "sumif":
                item = criteria[0]
                formula = f"=SUMIF({_analysis_range(record, str(item['column']))},{_criterion_formula_value(item)},{_analysis_range(record, sum_column)})"
            else:
                arguments = [_analysis_range(record, sum_column)]
                for item in criteria:
                    arguments.extend([_analysis_range(record, str(item["column"])), _criterion_formula_value(item)])
                formula = f"=SUMIFS({','.join(arguments)})"
        columns = list(dict.fromkeys(columns))
        return finish(f"{operation_id.upper()} sonucu {value:,.2f}.", value, columns, formula)

    column, clarification = _resolve_required(record, command, mode, operation, "metric_column", "column", numeric=operation_id in {"sum", "average"})
    if clarification:
        return clarification
    if operation_id == "sum":
        value = float(_numeric_series(record, column).sum())
        formula = f"=SUM({_analysis_range(record, column)})"
    elif operation_id == "average":
        value = float(_numeric_series(record, column).mean())
        formula = f"=AVERAGE({_analysis_range(record, column)})"
    else:  # counta
        value = int(record.dataframe[column].notna().sum())
        formula = f"=COUNTA({_analysis_range(record, column)})"
    return finish(f"{column} sütununun {operation_id.upper()} sonucu {value:,.2f}.", value, [column], formula)


def _execute_formula(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    operation_id = operation["id"]
    source, clarification = _resolve_required(record, command, mode, operation, "source_column", "source_column", numeric=False)
    if clarification:
        return clarification
    target = operation.get("target_column") or f"{source}_{operation_id.upper()}"
    series = record.dataframe[source]
    formulas: list[str] = []
    values: list[Any] = []

    if operation_id == "if":
        operator = str(operation.get("operator") or ">")
        compare = operation.get("value", 0)
        true_value, false_value = operation.get("true_value", "Evet"), operation.get("false_value", "Hayır")
        mask = _criteria_mask(record.dataframe, [{"column": source, "operator": operator, "value": compare}])
        values = [true_value if item else false_value for item in mask]
        formulas = [f"=IF({_operator_formula(_column_ref(record, source, row), operator, compare)},{_excel_literal(true_value)},{_excel_literal(false_value)})" for row in range(2, len(series) + 2)]
    elif operation_id == "ifs":
        cases = operation.get("cases") or []
        if not cases:
            raise ExcelWizardError(400, "cases_required", "ÇOKEĞER için koşulları ve sonuçlarını belirtin.")
        default = operation.get("default", "")
        for offset, original in enumerate(series, 2):
            selected = default
            pieces: list[str] = []
            for case in cases:
                operator, compare, result = str(case.get("operator") or "="), case.get("value"), case.get("result")
                condition = bool(_criteria_mask(pd.DataFrame({source: [original]}), [{"column": source, "operator": operator, "value": compare}]).iloc[0])
                if condition and selected == default:
                    selected = result
                pieces.extend([_operator_formula(_column_ref(record, source, offset), operator, compare), _excel_literal(result)])
            values.append(selected)
            pieces.extend(["TRUE", _excel_literal(default)])
            formulas.append("=IFS(" + ",".join(pieces) + ")")
    elif operation_id in {"left", "right"}:
        count = max(1, int(operation.get("count") or 1))
        text = series.fillna("").astype(str)
        values = text.str[:count].tolist() if operation_id == "left" else text.str[-count:].tolist()
        formulas = [f"={operation_id.upper()}({_column_ref(record, source, row)},{count})" for row in range(2, len(series) + 2)]
    elif operation_id == "mid":
        start, count = max(1, int(operation.get("start") or 1)), max(1, int(operation.get("count") or 1))
        values = series.fillna("").astype(str).str.slice(start - 1, start - 1 + count).tolist()
        formulas = [f"=MID({_column_ref(record, source, row)},{start},{count})" for row in range(2, len(series) + 2)]
    elif operation_id in {"upper", "lower"}:
        text = series.fillna("").astype(str)
        values = (text.str.upper() if operation_id == "upper" else text.str.lower()).tolist()
        formulas = [f"={operation_id.upper()}({_column_ref(record, source, row)})" for row in range(2, len(series) + 2)]
    elif operation_id == "substitute":
        old, new = str(operation.get("old") or ""), str(operation.get("new") or "")
        values = series.fillna("").astype(str).str.replace(old, new, regex=False).tolist()
        formulas = [f"=SUBSTITUTE({_column_ref(record, source, row)},{_excel_literal(old)},{_excel_literal(new)})" for row in range(2, len(series) + 2)]
    elif operation_id == "find":
        needle = str(operation.get("text") or "")
        values = [str(value).find(needle) + 1 if needle in str(value) else None for value in series.fillna("")]
        formulas = [f"=IFERROR(FIND({_excel_literal(needle)},{_column_ref(record, source, row)}),\"\")" for row in range(2, len(series) + 2)]
    elif operation_id == "len":
        values = series.fillna("").astype(str).str.len().tolist()
        formulas = [f"=LEN({_column_ref(record, source, row)})" for row in range(2, len(series) + 2)]
    elif operation_id in {"day", "month", "year"}:
        parsed = pd.to_datetime(series, errors="coerce")
        values = getattr(parsed.dt, operation_id).where(parsed.notna(), None).tolist()
        formulas = [f"={operation_id.upper()}({_column_ref(record, source, row)})" for row in range(2, len(series) + 2)]
    else:
        raise ExcelWizardError(400, "unsupported_formula", f"{operation_id} formülü uygulanamadı.")

    _write_formula_column(record, str(target), formulas, values)
    return {"summary": f"{target} sütunu {operation_id.upper()} formülüyle oluşturuldu.", "metrics": {"rows_updated": len(values)}, "columns_used": [source, str(target)], "is_mutation": True}


def _execute_concat(record: WorkbookSession, operation: dict[str, Any]) -> dict[str, Any]:
    columns = operation.get("columns") or operation.get("source_columns") or []
    if isinstance(columns, str):
        columns = [columns]
    if operation.get("concat_column"):
        columns = list(dict.fromkeys([*columns, str(operation["concat_column"])]))
        operation["columns"] = columns
        operation.pop("concat_column", None)
    columns = [str(column) for column in columns]
    if any(column not in record.dataframe.columns for column in columns):
        raise ExcelWizardError(400, "columns_required", "Birleştirmek için en az iki geçerli sütun belirtin.")
    if len(columns) < 2:
        candidates = [str(column) for column in record.dataframe.columns if str(column) not in columns]
        return _clarification(record, "", "operations", operation, "concat_column", candidates[:8], "Birleştirilecek diğer sütunu seçin.")
    target = str(operation.get("target_column") or "Birleştirilmiş_Metin")
    delimiter = str(operation.get("delimiter") or (" " if operation["id"] == "concat" else ", "))
    values = record.dataframe[columns].fillna("").astype(str).agg(delimiter.join, axis=1).str.strip()
    formulas = []
    for row in range(2, len(record.dataframe) + 2):
        refs = [_column_ref(record, column, row) for column in columns]
        if operation["id"] == "textjoin":
            formulas.append(f"=TEXTJOIN({_excel_literal(delimiter)},TRUE,{','.join(refs)})")
        else:
            formulas.append("=CONCAT(" + f",{_excel_literal(delimiter)},".join(refs) + ")")
    _write_formula_column(record, target, formulas, values.tolist())
    return {"summary": f"{target} sütunu {len(columns)} sütun birleştirilerek oluşturuldu.", "metrics": {"rows_updated": len(values)}, "columns_used": columns + [target], "is_mutation": True}


def _execute_date(record: WorkbookSession, operation: dict[str, Any]) -> dict[str, Any]:
    operation_id = operation["id"]
    target = str(operation.get("target_column") or {"today": "Bugün", "now": "Şimdi", "date": "Tarih"}[operation_id])
    now = datetime.now()
    if operation_id == "today":
        values, formulas = [now.date()] * len(record.dataframe), ["=TODAY()"] * len(record.dataframe)
    elif operation_id == "now":
        values, formulas = [now] * len(record.dataframe), ["=NOW()"] * len(record.dataframe)
    else:
        year_column = operation.get("year_column")
        month_column = operation.get("month_column")
        day_column = operation.get("day_column")
        if all(column in record.dataframe.columns for column in (year_column, month_column, day_column)):
            parsed = pd.to_datetime({
                "year": pd.to_numeric(record.dataframe[year_column], errors="coerce"),
                "month": pd.to_numeric(record.dataframe[month_column], errors="coerce"),
                "day": pd.to_numeric(record.dataframe[day_column], errors="coerce"),
            }, errors="coerce")
            values = [value.date() if pd.notna(value) else None for value in parsed]
            formulas = [
                f"=DATE({_column_ref(record, year_column, row)},{_column_ref(record, month_column, row)},{_column_ref(record, day_column, row)})"
                for row in range(2, len(record.dataframe) + 2)
            ]
            _write_formula_column(record, target, formulas, values)
            return {"summary": f"{target} sütunu {year_column}, {month_column} ve {day_column} sütunlarından oluşturuldu.", "metrics": {"rows_updated": len(values)}, "columns_used": [year_column, month_column, day_column, target], "is_mutation": True}
        year = int(operation.get("year") or now.year)
        month = int(operation.get("month") or now.month)
        day = int(operation.get("day") or now.day)
        value = date(year, month, day)
        values, formulas = [value] * len(record.dataframe), [f"=DATE({year},{month},{day})"] * len(record.dataframe)
    _write_formula_column(record, target, formulas, values)
    return {"summary": f"{target} tarih sütunu oluşturuldu.", "metrics": {"rows_updated": len(values)}, "columns_used": [target], "is_mutation": True}


def _execute_lookup(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    source, clarification = _resolve_required(record, command, mode, operation, "lookup_key_column", "source_column")
    if clarification:
        return clarification
    if operation["id"] == "index_match" and operation.get("match_value") not in (None, "") and not operation.get("lookup_sheet"):
        match_value = operation["match_value"]
        mask = record.dataframe[source].astype(str).str.casefold() == str(match_value).casefold()
        position = int(mask.to_numpy().argmax()) + 1 if mask.any() else None
        return {
            "summary": f"{match_value}, {source} sütununda " + (f"{position}. veri satırında bulundu." if position else "bulunamadı."),
            "metrics": {"match_position": position},
            "columns_used": [source],
            "is_mutation": False,
        }
    lookup_sheet = str(operation.get("lookup_sheet") or "")
    if not lookup_sheet or lookup_sheet not in record.workbook.sheetnames:
        other_sheets = [sheet for sheet in record.workbook.sheetnames if sheet != record.active_sheet]
        return _clarification(record, command, mode, operation, "lookup_sheet", other_sheets, "Arama tablosu hangi sayfada?")
    lookup_ws = record.workbook[lookup_sheet]
    if operation["id"] == "hlookup":
        return_row = max(2, int(operation.get("return_row") or 2))
        if return_row > lookup_ws.max_row:
            raise ExcelWizardError(400, "invalid_return_row", f"Arama sayfasında {return_row}. satır bulunmuyor.")
        horizontal_mapping: dict[Any, Any] = {}
        for column_index in range(1, lookup_ws.max_column + 1):
            key = lookup_ws.cell(1, column_index).value
            if key not in (None, "") and key not in horizontal_mapping:
                horizontal_mapping[key] = lookup_ws.cell(return_row, column_index).value
        if not horizontal_mapping:
            raise ExcelWizardError(400, "invalid_lookup_table", "YATAYARA sayfasının ilk satırında arama anahtarı bulunamadı.")
        target = str(operation.get("target_column") or f"{lookup_sheet}_YATAYARA")
        normalized_mapping = {str(key).casefold(): value for key, value in horizontal_mapping.items()}
        values = [normalized_mapping.get(str(value).casefold()) for value in record.dataframe[source]]
        source_letter = get_column_letter(list(record.dataframe.columns).index(source) + 1)
        sheet = lookup_sheet.replace("'", "''")
        formulas = [
            f"=IFERROR(HLOOKUP({source_letter}{row},'{sheet}'!$1:${return_row},{return_row},FALSE),\"\")"
            for row in range(2, len(record.dataframe) + 2)
        ]
        _write_formula_column(record, target, formulas, values)
        return {
            "summary": f"{target} sütunu {lookup_sheet} sayfasının {return_row}. satırından HLOOKUP ile eşleştirildi.",
            "metrics": {"matched_rows": int(pd.Series(values).notna().sum())},
            "columns_used": [source, f"{lookup_sheet}!1:{return_row}", target],
            "is_mutation": True,
        }
    lookup_df = _dataframe_from_worksheet(lookup_ws)
    key_column = str(operation.get("lookup_key_column") or lookup_df.columns[0] if len(lookup_df.columns) else "")
    return_column = str(operation.get("return_column") or "")
    if key_column not in lookup_df.columns or return_column not in lookup_df.columns:
        options = [str(column) for column in lookup_df.columns]
        return _clarification(record, command, mode, operation, "return_column", options, "Arama sayfasından hangi sütun getirilsin?")
    target = str(operation.get("target_column") or return_column)
    mapping = lookup_df.drop_duplicates(key_column).set_index(key_column)[return_column]
    values = record.dataframe[source].map(mapping).tolist()
    source_letter = get_column_letter(list(record.dataframe.columns).index(source) + 1)
    key_index = list(lookup_df.columns).index(key_column) + 1
    return_index = list(lookup_df.columns).index(return_column) + 1
    key_letter, return_letter = get_column_letter(key_index), get_column_letter(return_index)
    formulas = []
    for row in range(2, len(record.dataframe) + 2):
        ref = f"{source_letter}{row}"
        sheet = lookup_sheet.replace("'", "''")
        if operation["id"] == "vlookup":
            if return_index >= key_index:
                index = return_index - key_index + 1
                formulas.append(f"=IFERROR(VLOOKUP({ref},'{sheet}'!${key_letter}:${return_letter},{index},FALSE),\"\")")
            else:
                # CHOOSE keeps VLOOKUP exact while also supporting a return
                # column positioned to the left of the key column.
                formulas.append(
                    f"=IFERROR(VLOOKUP({ref},CHOOSE({{1,2}},'{sheet}'!${key_letter}:${key_letter},"
                    f"'{sheet}'!${return_letter}:${return_letter}),2,FALSE),\"\")"
                )
        elif operation["id"] == "index_match":
            formulas.append(f"=IFERROR(INDEX('{sheet}'!${return_letter}:${return_letter},MATCH({ref},'{sheet}'!${key_letter}:${key_letter},0)),\"\")")
        elif operation["id"] == "xlookup":
            formulas.append(f"=XLOOKUP({ref},'{sheet}'!${key_letter}:${key_letter},'{sheet}'!${return_letter}:${return_letter},\"\")")
        elif operation["id"] == "lookup":
            # LOOKUP's ordinary form is approximate and can disagree with the
            # exact mapping calculated above when SKU keys are unsorted.  This
            # classic LOOKUP form performs an exact last-match lookup instead.
            formulas.append(f"=IFERROR(LOOKUP(2,1/('{sheet}'!${key_letter}:${key_letter}={ref}),'{sheet}'!${return_letter}:${return_letter}),\"\")")
    _write_formula_column(record, target, formulas, values)
    return {"summary": f"{target} sütunu {lookup_sheet} sayfasından {operation['id'].upper()} ile eşleştirildi.", "metrics": {"matched_rows": int(pd.Series(values).notna().sum())}, "columns_used": [source, key_column, return_column, target], "is_mutation": True}


def _execute_data_tool(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    operation_id = operation["id"]
    if operation_id == "filter":
        criteria = operation.get("criteria") or []
        if not criteria:
            raise ExcelWizardError(400, "criteria_required", "Filtre için koşul belirtin.")
        mask = _criteria_mask(record.dataframe, criteria)
        matching, removed = int(mask.sum()), int((~mask).sum())
        record.dataframe = record.dataframe.loc[mask].reset_index(drop=True)
        _sync_values(record)
        return {"summary": f"Filtre uygulandı; {matching:,} satır eşleşti.", "metrics": {"matching_rows": matching, "removed_rows": removed}, "columns_used": list(dict.fromkeys(str(item["column"]) for item in criteria)), "is_mutation": True}
    column, clarification = _resolve_required(record, command, mode, operation, "target_column", "column")
    if clarification:
        return clarification
    worksheet = record.workbook[record.active_sheet]
    column_index = list(record.dataframe.columns).index(column) + 1
    letter = get_column_letter(column_index)
    if operation_id == "sort":
        ascending = bool(operation.get("ascending", True))
        record.dataframe = record.dataframe.sort_values(column, ascending=ascending, na_position="last", kind="mergesort").reset_index(drop=True)
        _sync_values(record)
        return {"summary": f"{column} sütunu {'artan' if ascending else 'azalan'} sıralandı.", "metrics": {"rows_sorted": len(record.dataframe)}, "columns_used": [column], "is_mutation": True}
    if operation_id == "conditional_format":
        operator = str(operation.get("operator") or "greaterThan")
        value = operation.get("value")
        fill = PatternFill("solid", fgColor=str(operation.get("color") or "FECACA").replace("#", ""))
        cell_range = f"{letter}2:{letter}{len(record.dataframe) + 1}"
        if operation.get("threshold") == "mean":
            average = float(_numeric_series(record, column).mean())
            worksheet.conditional_formatting.add(cell_range, CellIsRule(operator="greaterThanOrEqual", formula=[str(average)], fill=PatternFill("solid", fgColor="DCFCE7")))
            worksheet.conditional_formatting.add(cell_range, CellIsRule(operator="lessThan", formula=[str(average)], fill=PatternFill("solid", fgColor="FEE2E2")))
        elif value is None:
            worksheet.conditional_formatting.add(cell_range, ColorScaleRule(start_type="min", start_color="FEE2E2", mid_type="percentile", mid_value=50, mid_color="FEF3C7", end_type="max", end_color="DCFCE7"))
        else:
            operator = operator if operator in {"greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "equal", "notEqual"} else "greaterThan"
            worksheet.conditional_formatting.add(cell_range, CellIsRule(operator=operator, formula=[str(value)], fill=fill))
        return {"summary": f"{column} sütununa koşullu biçimlendirme eklendi.", "metrics": {"formatted_rows": len(record.dataframe)}, "columns_used": [column], "is_mutation": True}
    if operation_id == "data_validation":
        values = [str(value) for value in (operation.get("values") or [])]
        if not values:
            raise ExcelWizardError(400, "values_required", "Veri doğrulama için izin verilen değerleri belirtin.")
        escaped = [value.replace('"', '""') for value in values]
        formula = '"' + ",".join(escaped) + '"'
        if len(formula) > 255:
            hidden_name = "_DataProvido_Validation"
            if hidden_name in record.workbook.sheetnames:
                del record.workbook[hidden_name]
            hidden = record.workbook.create_sheet(hidden_name)
            for index, value in enumerate(values, 1):
                hidden.cell(index, 1, value)
            hidden.sheet_state = "hidden"
            formula = f"'{hidden_name}'!$A$1:$A${len(values)}"
        validation = DataValidation(type="list", formula1=formula, allow_blank=True)
        validation.error = "Listeden geçerli bir değer seçin."
        validation.errorTitle = "Geçersiz değer"
        worksheet.add_data_validation(validation)
        validation.add(f"{letter}2:{letter}{max(2, len(record.dataframe) + 1)}")
        return {"summary": f"{column} sütununa {len(values)} seçenekli veri doğrulama eklendi.", "metrics": {"allowed_values": len(values)}, "columns_used": [column], "is_mutation": True}
    if operation_id == "text_to_columns":
        delimiter = str(operation.get("delimiter") or ",")
        split = record.dataframe[column].fillna("").astype(str).str.split(delimiter, expand=True)
        targets = operation.get("target_columns") or [f"{column}_{index + 1}" for index in range(split.shape[1])]
        if len(targets) != split.shape[1]:
            raise ExcelWizardError(400, "target_columns_mismatch", "Hedef sütun sayısı ayrıştırılan parça sayısıyla eşleşmiyor.")
        for index, target in enumerate(targets):
            record.dataframe[str(target)] = split[index]
        _sync_values(record)
        return {"summary": f"{column} sütunu {split.shape[1]} sütuna ayrıldı.", "metrics": {"created_columns": split.shape[1]}, "columns_used": [column] + [str(item) for item in targets], "is_mutation": True}
    raise ExcelWizardError(400, "unsupported_operation", f"{operation_id} desteklenmiyor.")


def _execute_pivot(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    index = operation.get("index") or operation.get("group_column")
    values = operation.get("values") or operation.get("metric_column")
    if not index:
        index, clarification = _resolve_required(record, command, mode, operation, operation.get("group_role") or "brand", "index")
        if clarification:
            return clarification
    if not values:
        values, clarification = _resolve_required(record, command, mode, operation, "metric_column", "values", numeric=True)
        if clarification:
            return clarification
    if isinstance(index, list):
        index = index[0] if index else ""
    if isinstance(values, list):
        values = values[0] if values else ""
    index, values = str(index), str(values)
    if index not in record.dataframe.columns or values not in record.dataframe.columns:
        raise ExcelWizardError(400, "invalid_column", "Pivot sütunlarından biri bulunamadı.")
    aggfunc = str(operation.get("aggfunc") or "sum").lower()
    if aggfunc not in {"sum", "mean", "count", "min", "max"}:
        raise ExcelWizardError(400, "invalid_aggregation", "Pivot hesaplaması sum, mean, count, min veya max olmalıdır.")
    columns = operation.get("columns")
    if isinstance(columns, list):
        columns = columns[0] if columns else None
    if columns:
        columns = str(columns)
        if columns not in record.dataframe.columns:
            raise ExcelWizardError(400, "invalid_column", f"'{columns}' pivot sütunu bulunamadı.")
        pivot = pd.pivot_table(record.dataframe, index=index, columns=columns, values=values, aggfunc=aggfunc, fill_value=0).reset_index()
        pivot.columns = [" | ".join(str(part) for part in column if str(part) not in ("", "None")) if isinstance(column, tuple) else str(column) for column in pivot.columns]
    else:
        pivot = record.dataframe.groupby(index, dropna=False)[values].agg(aggfunc).reset_index().sort_values(values, ascending=False)
    sheet_name = str(operation.get("sheet_name") or "Pivot_Analysis")[:31]
    if sheet_name in record.workbook.sheetnames:
        del record.workbook[sheet_name]
    worksheet = record.workbook.create_sheet(sheet_name)
    for col_index, column in enumerate(pivot.columns, 1):
        cell = worksheet.cell(1, col_index, str(column))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row_index, row in enumerate(pivot.itertuples(index=False, name=None), 2):
        for col_index, value in enumerate(row, 1):
            worksheet.cell(row_index, col_index, _json_value(value))
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    used_columns = [index] + ([columns] if columns else []) + [values]
    return {"summary": f"{index} bazında {values} için {aggfunc} pivot raporu '{sheet_name}' sayfasına eklendi.", "metrics": {"groups": len(pivot)}, "columns_used": used_columns, "preview_rows": dataframe_preview(pivot)["rows"], "is_mutation": True}


def _execute_chart(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    category = operation.get("category_column")
    values = operation.get("value_columns") or ([operation.get("value_column")] if operation.get("value_column") else [])
    if not category:
        category, clarification = _resolve_required(record, command, mode, operation, operation.get("group_role") or "category", "category_column")
        if clarification:
            return clarification
    if not values:
        value, clarification = _resolve_required(record, command, mode, operation, "metric_column", "value_column", numeric=True)
        if clarification:
            return clarification
        values = [value]
    values = [str(value) for value in values]
    if str(category) not in record.dataframe.columns or any(value not in record.dataframe.columns for value in values):
        raise ExcelWizardError(400, "invalid_column", "Grafik sütunlarından biri bulunamadı.")
    worksheet = record.workbook[record.active_sheet]
    chart_type = str(operation.get("chart_type") or "bar").lower()
    chart = PieChart() if chart_type in {"pie", "pasta"} else (LineChart() if chart_type in {"line", "cizgi"} else BarChart())
    chart_values = values[:1] if chart_type in {"pie", "pasta"} else values
    chart_frame = pd.DataFrame({str(category): record.dataframe[str(category)]})
    for value in chart_values:
        chart_frame[value] = _numeric_series(record, value)
    chart_frame = chart_frame.groupby(str(category), dropna=False)[chart_values].sum(min_count=1).reset_index()
    if chart_type not in {"line", "cizgi"}:
        chart_frame = chart_frame.sort_values(chart_values[0], ascending=False, na_position="last")

    base_sheet_name = str(operation.get("data_sheet_name") or "Chart_Analysis")[:31]
    chart_sheet_name = base_sheet_name
    suffix = 2
    while chart_sheet_name in record.workbook.sheetnames:
        tail = f"_{suffix}"
        chart_sheet_name = base_sheet_name[:31 - len(tail)] + tail
        suffix += 1
    chart_sheet = record.workbook.create_sheet(chart_sheet_name)
    for column_index, column in enumerate(chart_frame.columns, 1):
        cell = chart_sheet.cell(1, column_index, str(column))
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        chart_sheet.column_dimensions[get_column_letter(column_index)].width = 22
    for row_index, row in enumerate(chart_frame.itertuples(index=False, name=None), 2):
        for column_index, value in enumerate(row, 1):
            chart_sheet.cell(row_index, column_index, _json_value(value))
    chart_sheet.freeze_panes = "A2"
    chart_sheet.auto_filter.ref = chart_sheet.dimensions

    for value_index in range(2, len(chart_values) + 2):
        chart.add_data(Reference(chart_sheet, min_col=value_index, min_row=1, max_row=len(chart_frame) + 1), titles_from_data=True)
    chart.set_categories(Reference(chart_sheet, min_col=1, min_row=2, max_row=len(chart_frame) + 1))
    title = str(operation.get("title") or f"{', '.join(values)} by {category}")
    chart.title = title
    chart.style = 10
    chart.height, chart.width = 8, 14
    if isinstance(chart, BarChart):
        chart.type = "col"
        chart.y_axis.title = ", ".join(chart_values)
        chart.x_axis.title = str(category)
    elif isinstance(chart, LineChart):
        chart.y_axis.title = ", ".join(chart_values)
        chart.x_axis.title = str(category)
    anchor = f"{get_column_letter(len(record.dataframe.columns) + 3)}2"
    worksheet.add_chart(chart, anchor)
    return {
        "summary": f"{title} grafiği eklendi; gruplanmış kaynak verisi '{chart_sheet_name}' sayfasında hazırlandı.",
        "metrics": {"series": len(chart_values), "categories": len(chart_frame)},
        "columns_used": [str(category)] + chart_values,
        "preview_rows": dataframe_preview(chart_frame)["rows"],
        "is_mutation": True,
    }


def _resolve_optional_role(
    record: WorkbookSession,
    role: str,
    param: str,
    *,
    numeric: bool = False,
) -> Optional[str]:
    """Resolve an enrichment column without blocking the core analysis.

    Optional metrics are included only when the heading is unambiguous.  A
    required metric still goes through ``_resolve_required`` and asks the user
    an actionable question when more than one column could be correct.
    """
    requested = record.role_mappings.get(param)
    resolved, candidates = resolve_column(record, role, requested=requested, numeric=numeric)
    if resolved and not candidates:
        record.role_mappings[role] = resolved
        record.role_mappings[param] = resolved
        return resolved
    return None


def _require_distinct_role(
    record: WorkbookSession,
    command: str,
    mode: str,
    operation: dict[str, Any],
    *,
    role: str,
    param: str,
    used: Iterable[str],
) -> tuple[Optional[str], Optional[dict[str, Any]]]:
    column, clarification = _resolve_required(record, command, mode, operation, role, param, numeric=True)
    if clarification or column not in set(used):
        return column, clarification
    candidates = [
        str(item) for item in record.dataframe.columns
        if str(item) not in set(used) and _is_numeric_column(record.dataframe, str(item))
    ][:8]
    return None, _clarification(
        record,
        command,
        mode,
        operation,
        param,
        candidates,
        f"{role.replace('_', ' ').title()} için diğer metriklerden farklı olan sütunu seçin.",
    )


def _execute_commerce(record: WorkbookSession, operation: dict[str, Any], command: str, mode: str) -> dict[str, Any]:
    analysis_type = str(operation.get("analysis_type") or "ranking").lower()
    question = _norm(command)
    if analysis_type == "ranking":
        if any(word in question for word in ["funnel", "pdp", "add to cart", "sepete"]):
            analysis_type = "funnel"
        elif any(word in question for word in ["stok", "stock", "inventory"]):
            analysis_type = "stock"
        elif any(word in question for word in ["kategori", "category"]):
            analysis_type = "category"
        elif any(word in question for word in ["marka", "brand"]):
            analysis_type = "brand"

    default_group_role = "category" if analysis_type == "category" else ("brand" if analysis_type == "brand" else "product")
    group_role = str(operation.get("group_role") or default_group_role)
    group, clarification = _resolve_required(record, command, mode, operation, group_role, "group_column")
    if clarification:
        return clarification

    if analysis_type == "funnel":
        views, clarification = _require_distinct_role(
            record, command, mode, operation, role="views", param="views_column", used=[]
        )
        if clarification:
            return clarification
        carts, clarification = _require_distinct_role(
            record, command, mode, operation, role="add_to_cart", param="add_to_cart_column", used=[views]
        )
        if clarification:
            return clarification
        transactions, clarification = _require_distinct_role(
            record, command, mode, operation, role="transactions", param="transactions_column", used=[views, carts]
        )
        if clarification:
            return clarification
        working = pd.DataFrame({
            group: record.dataframe[group],
            views: _numeric_series(record, views),
            carts: _numeric_series(record, carts),
            transactions: _numeric_series(record, transactions),
        })
        grouped = working.groupby(group, dropna=False)[[views, carts, transactions]].sum(min_count=1).reset_index()
        grouped["view_to_cart_rate"] = grouped[carts].div(grouped[views].replace(0, pd.NA))
        grouped["cart_to_purchase_rate"] = grouped[transactions].div(grouped[carts].replace(0, pd.NA))
        grouped["view_to_purchase_rate"] = grouped[transactions].div(grouped[views].replace(0, pd.NA))
        grouped["view_to_cart_drop_off"] = grouped[views] - grouped[carts]
        grouped["cart_to_purchase_drop_off"] = grouped[carts] - grouped[transactions]
        grouped["lost_purchases"] = grouped[views] - grouped[transactions]
        grouped["data_quality_issue"] = (grouped[carts] > grouped[views]) | (grouped[transactions] > grouped[carts])
        grouped = grouped.sort_values(["lost_purchases", "view_to_purchase_rate"], ascending=[False, True], na_position="last")
        grouped.insert(0, "rank", range(1, len(grouped) + 1))

        total_views = float(grouped[views].sum())
        total_carts = float(grouped[carts].sum())
        total_transactions = float(grouped[transactions].sum())
        stage_losses = {
            "view_to_cart": total_views - total_carts,
            "cart_to_purchase": total_carts - total_transactions,
        }
        largest_stage = max(stage_losses, key=stage_losses.get) if stage_losses else None
        valid_conversion = grouped.loc[grouped[views] > 0]
        weakest = valid_conversion.sort_values("view_to_purchase_rate", na_position="last").iloc[0] if not valid_conversion.empty else None
        metrics = {
            "total_views": total_views,
            "total_add_to_carts": total_carts,
            "total_transactions": total_transactions,
            "view_to_cart_rate": total_carts / total_views if total_views else None,
            "cart_to_purchase_rate": total_transactions / total_carts if total_carts else None,
            "view_to_purchase_rate": total_transactions / total_views if total_views else None,
            "largest_drop_stage": largest_stage,
            "largest_drop_count": stage_losses.get(largest_stage) if largest_stage else None,
            "weakest_product": _json_value(weakest[group]) if weakest is not None else None,
            "data_quality_issues": int(grouped["data_quality_issue"].sum()),
        }
        stage_label = "görüntülemeden sepete" if largest_stage == "view_to_cart" else "sepetten satın almaya"
        summary = (
            f"En büyük toplam kayıp {stage_label} aşamasında {stage_losses.get(largest_stage, 0):,.0f}. "
            f"{total_views:,.0f} görüntüleme, {total_carts:,.0f} sepete ekleme ve {total_transactions:,.0f} satın alma analiz edildi."
        )
        return {
            "summary": summary,
            "metrics": metrics,
            "columns_used": [group, views, carts, transactions],
            "preview_rows": dataframe_preview(grouped.head(min(100, max(1, int(operation.get("top_n") or 10)))))["rows"],
            "is_mutation": False,
        }

    metric, clarification = _resolve_required(record, command, mode, operation, "revenue", "metric_column", numeric=True)
    if clarification:
        return clarification
    metric_values = _numeric_series(record, metric)

    if analysis_type == "stock":
        stock, clarification = _resolve_required(record, command, mode, operation, "stock", "stock_column", numeric=True)
        if clarification:
            return clarification
        demand = (
            _resolve_optional_role(record, "units", "units_column", numeric=True)
            or _resolve_optional_role(record, "transactions", "transactions_column", numeric=True)
            or _resolve_optional_role(record, "orders", "orders_column", numeric=True)
            or _resolve_optional_role(record, "views", "views_column", numeric=True)
        )
        columns = [metric, stock] + ([demand] if demand and demand not in {metric, stock} else [])
        working = pd.DataFrame({group: record.dataframe[group]})
        for column in columns:
            working[column] = _numeric_series(record, column)
        grouped = working.groupby(group, dropna=False)[columns].sum(min_count=1).reset_index()
        demand_column = demand if demand and demand not in {metric, stock} else metric
        positive_stock = grouped.loc[grouped[stock] > 0, stock]
        low_stock_threshold = float(positive_stock.quantile(0.25)) if not positive_stock.empty else 0.0
        median_stock = float(positive_stock.median()) if not positive_stock.empty else 0.0
        median_demand = float(grouped[demand_column].median()) if not grouped.empty else 0.0
        median_revenue = float(grouped[metric].median()) if not grouped.empty else 0.0
        grouped["revenue_per_stock"] = grouped[metric].div(grouped[stock].clip(lower=1))
        grouped["demand_per_stock"] = grouped[demand_column].div(grouped[stock].clip(lower=1))
        grouped["priority_score"] = (
            grouped[metric].rank(pct=True, method="average") * 0.45
            + grouped[demand_column].rank(pct=True, method="average") * 0.35
            + grouped[stock].rank(pct=True, ascending=False, method="average") * 0.20
        ) * 100

        def risk_label(row: pd.Series) -> str:
            if row[stock] <= 0 and (row[demand_column] > 0 or row[metric] > 0):
                return "Critical"
            if row[stock] <= low_stock_threshold and (row[demand_column] >= median_demand or row[metric] >= median_revenue):
                return "High"
            if row[stock] <= median_stock and row[demand_column] >= median_demand:
                return "Medium"
            return "Low"

        grouped["stock_risk"] = grouped.apply(risk_label, axis=1)
        risk_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
        grouped["_risk_order"] = grouped["stock_risk"].map(risk_order)
        grouped = grouped.sort_values(["_risk_order", "priority_score"], ascending=[True, False]).drop(columns=["_risk_order"])
        grouped.insert(0, "rank", range(1, len(grouped) + 1))
        priority = grouped.iloc[0] if not grouped.empty else None
        at_risk = grouped["stock_risk"].isin(["Critical", "High"])
        metrics = {
            "total_revenue": float(grouped[metric].sum()),
            "total_stock": float(grouped[stock].sum()),
            "at_risk_products": int(at_risk.sum()),
            "critical_products": int((grouped["stock_risk"] == "Critical").sum()),
            "high_risk_products": int((grouped["stock_risk"] == "High").sum()),
            "revenue_at_risk": float(grouped.loc[at_risk, metric].sum()),
            "priority_product": _json_value(priority[group]) if priority is not None else None,
            "low_stock_threshold": low_stock_threshold,
            "demand_metric": demand_column,
        }
        summary = (
            f"{int(at_risk.sum())} ürün kritik veya yüksek stok riski taşıyor."
            + (f" Gelir ve talep etkisine göre ilk öncelik {priority[group]}." if priority is not None else "")
        )
        return {
            "summary": summary,
            "metrics": metrics,
            "columns_used": [group, metric, stock] + ([demand] if demand and demand not in {metric, stock} else []),
            "preview_rows": dataframe_preview(grouped.head(min(100, max(1, int(operation.get("top_n") or 10)))))["rows"],
            "is_mutation": False,
        }

    working = pd.DataFrame({group: record.dataframe[group], metric: metric_values})
    order_column: Optional[str] = None
    units_column: Optional[str] = None
    if analysis_type == "category":
        order_column = (
            _resolve_optional_role(record, "order", "order_column", numeric=False)
            or _resolve_optional_role(record, "orders", "orders_column", numeric=True)
            or _resolve_optional_role(record, "transactions", "transactions_column", numeric=True)
        )
        units_column = _resolve_optional_role(record, "units", "units_column", numeric=True)
        aggregations: dict[str, Any] = {metric: "sum"}
        if units_column and units_column not in aggregations:
            working[units_column] = _numeric_series(record, units_column)
            aggregations[units_column] = "sum"
        if order_column and order_column not in aggregations:
            working[order_column] = record.dataframe[order_column]
            normalized_order = _norm(order_column)
            order_is_count = _is_numeric_column(record.dataframe, order_column) and any(
                token in normalized_order.split() for token in ["count", "quantity", "qty", "sayisi", "adedi", "orders", "transactions"]
            )
            aggregations[order_column] = "sum" if order_is_count else pd.Series.nunique
        grouped = working.groupby(group, dropna=False).agg(aggregations).reset_index()
        if order_column:
            grouped["average_order_value"] = grouped[metric].div(pd.to_numeric(grouped[order_column], errors="coerce").replace(0, pd.NA))
    else:
        grouped = working.groupby(group, dropna=False)[metric].sum(min_count=1).reset_index()

    grouped = grouped.sort_values(metric, ascending=False)
    total = float(grouped[metric].sum())
    grouped["share"] = grouped[metric].div(total) if total else pd.NA
    grouped["gap_to_leader"] = float(grouped.iloc[0][metric]) - grouped[metric] if not grouped.empty else pd.Series(dtype=float)
    grouped.insert(0, "rank", range(1, len(grouped) + 1))
    top_n = min(100, max(1, int(operation.get("top_n") or 10)))
    top = grouped.head(top_n)
    leader = grouped.iloc[0] if not grouped.empty else None
    runner_up = grouped.iloc[1] if len(grouped) > 1 else None
    weakest = grouped.iloc[-1] if not grouped.empty else None
    metrics: dict[str, Any] = {
        "total": total,
        "groups": int(len(grouped)),
        "average_per_group": float(grouped[metric].mean()) if not grouped.empty else None,
        "top_3_share": float(grouped.head(3)[metric].sum()) / total if total else None,
        "leader": _json_value(leader[group]) if leader is not None else None,
        "leader_value": float(leader[metric]) if leader is not None else None,
        "leader_share": float(leader[metric]) / total if leader is not None and total else None,
        "runner_up": _json_value(runner_up[group]) if runner_up is not None else None,
        "runner_up_value": float(runner_up[metric]) if runner_up is not None else None,
        "leader_gap": float(leader[metric] - runner_up[metric]) if leader is not None and runner_up is not None else None,
    }
    if analysis_type == "category":
        metrics.update({
            "total_revenue": total,
            "total_orders": float(pd.to_numeric(grouped[order_column], errors="coerce").sum()) if order_column else None,
            "total_units": float(pd.to_numeric(grouped[units_column], errors="coerce").sum()) if units_column else None,
            "average_order_value": (
                total / float(pd.to_numeric(grouped[order_column], errors="coerce").sum())
                if order_column and float(pd.to_numeric(grouped[order_column], errors="coerce").sum())
                else None
            ),
            "strongest_category": _json_value(leader[group]) if leader is not None else None,
            "weakest_category": _json_value(weakest[group]) if weakest is not None else None,
        })
        summary = (
            f"En güçlü kategori {leader[group]} ve {metric} değeri {float(leader[metric]):,.2f}. "
            f"En zayıf kategori {weakest[group]} ve toplam {metric} {total:,.2f}."
            if leader is not None and weakest is not None else f"{metric} için analiz edilebilir kategori bulunamadı."
        )
    elif leader is not None:
        gap_text = f" İkinci sıradaki {runner_up[group]} ile fark {float(leader[metric] - runner_up[metric]):,.2f}." if runner_up is not None else ""
        summary = f"{group} bazında en yüksek {metric}, {leader[group]} için {float(leader[metric]):,.2f}. Toplam {metric}: {total:,.2f}.{gap_text}"
    else:
        summary = f"{metric} için analiz edilebilir satır bulunamadı."
    return {
        "summary": summary,
        "metrics": metrics,
        "columns_used": [group, metric] + ([order_column] if order_column else []) + ([units_column] if units_column else []),
        "preview_rows": dataframe_preview(top)["rows"],
        "is_mutation": False,
    }


def _apply_clarification(record: WorkbookSession, operation: dict[str, Any], answers: dict[str, Any]) -> dict[str, Any]:
    clarification_id = str(answers.get("clarification_id") or "")
    pending = record.pending.pop(clarification_id, None) if clarification_id else None
    if pending:
        operation = dict(pending.operation)
        selected = answers.get("selected_value", answers.get("answer"))
        if selected not in (None, ""):
            operation[pending.role] = selected
            record.role_mappings[pending.role] = str(selected)
    allowed = {str(column) for column in record.dataframe.columns}
    for role, chosen in answers.items():
        if role in {"clarification_id", "selected_value", "answer"} or chosen in (None, ""):
            continue
        if role == "lookup_sheet":
            if str(chosen) not in record.workbook.sheetnames:
                raise ExcelWizardError(400, "invalid_sheet", "Seçilen çalışma sayfası bulunamadı.")
            operation[role] = str(chosen)
            continue
        if isinstance(chosen, list):
            if any(str(item) not in allowed for item in chosen):
                raise ExcelWizardError(400, "invalid_column", "Seçilen sütunlardan biri çalışma kitabında yok.")
            operation[role] = [str(item) for item in chosen]
            continue
        if str(chosen) not in allowed and role.endswith(("column", "columns")):
            raise ExcelWizardError(400, "invalid_column", "Seçilen sütun çalışma kitabında yok.")
        operation[role] = str(chosen)
        if role in ROLE_ALIASES or role.endswith("_column"):
            record.role_mappings[role] = str(chosen)
    return operation


def execute_operation(
    record: WorkbookSession,
    command: str,
    mode: str,
    operation: Optional[dict[str, Any]] = None,
    clarification_answers: Optional[dict[str, Any]] = None,
    planner: Optional[Callable[[str, list[dict[str, Any]], list[str]], Any]] = None,
    operation_hint: str = "",
) -> dict[str, Any]:
    if mode not in {"operations", "commerce"}:
        raise ExcelWizardError(422, "invalid_mode", "mode yalnızca 'operations' veya 'commerce' olabilir.")
    command = str(command or "").strip()
    if not command and not operation:
        raise ExcelWizardError(422, "command_required", "Yapılacak işlemi yazın.")
    planner_source = "structured" if operation else "deterministic"
    plan = dict(operation or {})
    if not plan:
        if operation_hint:
            plan = build_operation_from_command(command, mode, record, operation_hint)
        else:
            ai_plan = _strict_ai_plan(planner, command, record)
            if ai_plan:
                plan, planner_source = ai_plan, "ai_schema_only"
            else:
                plan = _infer_operation(command, mode, record)
    if plan.get("_batch"):
        return execute_operations(
            record,
            command,
            mode,
            list(plan["_batch"]),
            clarification_answers=clarification_answers,
            planner_source="deterministic",
        )
    operation_id = str(plan.get("id") or plan.get("type") or "").strip().lower()
    if operation_id not in CAPABILITY_IDS:
        raise ExcelWizardError(400, "unsupported_operation", "İstenen işlem desteklenen Excel işlemleri arasında değil.")
    plan["id"] = operation_id
    if clarification_answers:
        plan = _apply_clarification(record, plan, clarification_answers)

    aggregate_ids = {"sum", "average", "sumproduct", "sumif", "sumifs", "countif", "countifs", "counta"}
    formula_ids = {"if", "ifs", "left", "right", "mid", "upper", "lower", "substitute", "find", "len", "day", "month", "year"}
    lookup_ids = {"vlookup", "hlookup", "index_match", "lookup", "xlookup"}
    mutation_ids = formula_ids | lookup_ids | {
        "concat", "textjoin", "today", "now", "date", "conditional_format",
        "filter", "sort", "data_validation", "text_to_columns", "pivot", "chart",
    }
    if mode != "operations" and operation_id in mutation_ids:
        raise ExcelWizardError(409, "read_only_mode", "Bu komut çalışma kitabını değiştirir. 'Excel işlemleri' modunu seçin.")
    if operation_id in aggregate_ids:
        outcome = _execute_aggregate(record, plan, command, mode)
    elif operation_id in formula_ids:
        outcome = _execute_formula(record, plan, command, mode)
    elif operation_id in {"concat", "textjoin"}:
        outcome = _execute_concat(record, plan)
    elif operation_id in {"today", "now", "date"}:
        outcome = _execute_date(record, plan)
    elif operation_id in lookup_ids:
        outcome = _execute_lookup(record, plan, command, mode)
    elif operation_id in {"conditional_format", "filter", "sort", "data_validation", "text_to_columns"}:
        outcome = _execute_data_tool(record, plan, command, mode)
    elif operation_id == "pivot":
        outcome = _execute_pivot(record, plan, command, mode)
    elif operation_id == "chart":
        outcome = _execute_chart(record, plan, command, mode)
    elif operation_id == "commerce_analysis":
        outcome = _execute_commerce(record, plan, command, mode)
    else:
        raise ExcelWizardError(400, "unsupported_operation", f"{operation_id} işlemi uygulanamadı.")

    if outcome.get("status") == "needs_clarification":
        return outcome
    is_mutation = bool(outcome.get("is_mutation", False)) and mode == "operations"
    # Commerce mode never mutates. Formula/data operations are rejected instead of
    # silently modifying in a read-only request.
    if outcome.get("is_mutation") and mode != "operations":
        raise ExcelWizardError(409, "read_only_mode", "Bu komut çalışma kitabını değiştirir. 'Excel işlemleri' modunu seçin.")
    if is_mutation:
        record.modified = True
        record.operation_log.append({"at": datetime.now(timezone.utc).isoformat(), "operation": operation_id, "columns": outcome.get("columns_used", [])})
    preview = dataframe_preview(record.dataframe)
    result = {
        "summary": outcome.get("summary", "İşlem tamamlandı."),
        "steps": [
            f"İşlem planı: {operation_id}",
            "Sütunlar doğrulandı: " + (", ".join(outcome.get("columns_used", [])) or "sütun kullanılmadı"),
            "Hesaplama yerel ve deterministik olarak tamamlandı.",
        ],
        "metrics": {key: _json_value(value) for key, value in outcome.get("metrics", {}).items()},
        "columns_used": outcome.get("columns_used", []),
        "preview_rows": outcome.get("preview_rows", preview["rows"]),
        "is_mutation": is_mutation,
        "operation_id": operation_id,
        "planner_source": planner_source,
    }
    return {
        "status": "success",
        "result": result,
        "preview": preview,
        "download_url": f"/api/excel-wizard/download/{record.workbook_id}",
    }


def execute_operations(
    record: WorkbookSession,
    command: str,
    mode: str,
    operations: list[dict[str, Any]],
    clarification_answers: Optional[dict[str, Any]] = None,
    planner_source: str = "structured",
) -> dict[str, Any]:
    if not operations:
        raise ExcelWizardError(422, "operations_required", "En az bir işlem belirtin.")
    snapshot_buffer = BytesIO()
    record.workbook.save(snapshot_buffer)
    snapshot = {
        "workbook": snapshot_buffer.getvalue(),
        "dataframe": record.dataframe.copy(deep=True),
        "original_dataframe": record.original_dataframe.copy(deep=True),
        "active_sheet": record.active_sheet,
        "modified": record.modified,
        "role_mappings": dict(record.role_mappings),
        "operation_log": deepcopy(record.operation_log),
        "pending": dict(record.pending),
    }

    def restore(keep_pending: Optional[PendingClarification] = None) -> None:
        keep_vba = bool(getattr(record.workbook, "vba_archive", None))
        record.workbook = load_workbook(BytesIO(snapshot["workbook"]), data_only=False, keep_vba=keep_vba)
        record.dataframe = snapshot["dataframe"].copy(deep=True)
        record.original_dataframe = snapshot["original_dataframe"].copy(deep=True)
        record.active_sheet = snapshot["active_sheet"]
        record.modified = snapshot["modified"]
        record.role_mappings = dict(snapshot["role_mappings"])
        record.operation_log = deepcopy(snapshot["operation_log"])
        record.pending = dict(snapshot["pending"])
        if keep_pending is not None:
            record.pending[keep_pending.clarification_id] = keep_pending

    summaries, steps, metrics, columns = [], [], {}, []
    operation_ids: list[str] = []
    mutation = False
    try:
        for index, operation in enumerate(operations):
            response = execute_operation(record, command, mode, operation=operation, clarification_answers=clarification_answers if index == 0 else None)
            if response.get("status") == "needs_clarification":
                clarification = record.pending.get(str(response.get("clarification_id") or ""))
                restore(clarification)
                return response
            result = response["result"]
            operation_ids.append(result["operation_id"])
            summaries.append(result["summary"])
            steps.extend(result["steps"])
            metrics.update({f"{result['operation_id']}_{key}": value for key, value in result["metrics"].items()})
            columns.extend(result["columns_used"])
            mutation = mutation or result["is_mutation"]
    except Exception:
        restore()
        raise
    preview = dataframe_preview(record.dataframe)
    return {
        "status": "success",
        "result": {
            "summary": " ".join(summaries),
            "steps": steps,
            "metrics": metrics,
            "columns_used": list(dict.fromkeys(columns)),
            "preview_rows": preview["rows"],
            "is_mutation": mutation,
            "operation_id": "batch",
            "operation_ids": operation_ids,
            "planner_source": planner_source,
        },
        "preview": preview,
        "download_url": f"/api/excel-wizard/download/{record.workbook_id}",
    }


def workbook_bytes(record: WorkbookSession) -> bytes:
    output = BytesIO()
    record.workbook.save(output)
    contents = output.getvalue()
    if len(contents) > excel_workbook_store.max_upload_bytes * 2:
        raise ExcelWizardError(413, "result_too_large", "İşlenmiş çalışma kitabı indirme boyutu sınırını aşıyor.")
    return contents
