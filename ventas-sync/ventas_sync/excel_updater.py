"""Upsert Ventas Diarias rows in the La Cata workbook."""

from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .pdf_parser import TotalesGenerales

DEFAULT_UBICACIONES = ["La Cata LMF", "La Cata DF", "La Cata SJM"]

# Match existing Ventas Diarias formatting (RD$ amounts).
# FECHA is written as Spanish text so Mac Excel always shows it correctly
# (locale date formats like [$-1C0A] often display as "2026-07-20 0:00:00" on Mac).
CURRENCY_FORMAT = '"RD$"\\ #,##0.00'
CURRENCY_COLS = (3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)  # C–N

_DIAS = (
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado",
    "domingo",
)
_MESES = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)
_MES_LOOKUP = {nombre: idx + 1 for idx, nombre in enumerate(_MESES)}
_MES_LOOKUP.update(
    {
        "setiembre": 9,  # variant
    }
)


def format_fecha_es(value: date) -> str:
    """Match sheet display: 'martes 30 de junio de 2026'."""
    return f"{_DIAS[value.weekday()]} {value.day} de {_MESES[value.month - 1]} de {value.year}"


def _parse_fecha_es(text: str) -> date | None:
    """Parse Spanish long dates with or without accented characters."""
    import re
    import unicodedata

    def strip_accents(s: str) -> str:
        return "".join(
            ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn"
        )

    cleaned = " ".join(text.strip().lower().split())
    cleaned = strip_accents(cleaned)
    # lunes 20 de julio de 2026  OR  martes 30 de junio 2026
    match = re.match(
        r"^(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\s+"
        r"(\d{1,2})\s+de\s+"
        r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)"
        r"(?:\s+de)?\s+"
        r"(\d{4})$",
        cleaned,
    )
    if not match:
        return None
    day = int(match.group(2))
    month = _MES_LOOKUP[match.group(3)]
    year = int(match.group(4))
    return date(year, month, day)


@dataclass
class DailySaleRow:
    fecha: date
    ubicacion: str
    efectivo: float
    tarjeta_bruta: float
    transferencias: float
    notas_credito: float
    itbs_cobrado: float | None = None
    pedidos_ya: float | None = None


def totales_to_row(totales: TotalesGenerales, ubicacion: str) -> DailySaleRow:
    return DailySaleRow(
        fecha=totales.fecha.date(),
        ubicacion=ubicacion,
        efectivo=totales.efectivo,
        tarjeta_bruta=totales.tarjeta,
        transferencias=totales.otros_pagos,
        notas_credito=totales.nota_credito,
        itbs_cobrado=totales.itbis,
    )


def _as_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        from openpyxl.utils.datetime import from_excel

        return from_excel(value).date()
    if isinstance(value, str):
        text = value.strip()
        parsed = _parse_fecha_es(text)
        if parsed:
            return parsed
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def _find_row(ws: Worksheet, fecha: date, ubicacion: str) -> int | None:
    for row in range(2, ws.max_row + 1):
        row_date = _as_date(ws.cell(row, 1).value)
        row_ubic = ws.cell(row, 2).value
        if row_date == fecha and row_ubic == ubicacion:
            return row
    return None


def _rows_for_fecha(ws: Worksheet, fecha: date) -> dict[str, int]:
    found: dict[str, int] = {}
    for row in range(2, ws.max_row + 1):
        row_date = _as_date(ws.cell(row, 1).value)
        if row_date != fecha:
            continue
        ubic = ws.cell(row, 2).value
        if isinstance(ubic, str):
            found[ubic] = row
    return found


def _last_data_row(ws: Worksheet) -> int:
    last = 1
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 1).value is not None or ws.cell(row, 2).value is not None:
            last = row
    return last


def _chronological_insert_row(ws: Worksheet, fecha: date) -> int:
    """Row index where a new date block should start (1-based)."""
    for row in range(2, ws.max_row + 1):
        row_date = _as_date(ws.cell(row, 1).value)
        if row_date is None:
            # Hit the empty/pre-seeded zone — insert here
            return row
        if row_date > fecha:
            return row
    return _last_data_row(ws) + 1


def _apply_row_formats(ws: Worksheet, row: int) -> None:
    # Text date — no time component possible
    ws.cell(row, 1).number_format = "@"
    for col in CURRENCY_COLS:
        ws.cell(row, col).number_format = CURRENCY_FORMAT


def _write_fecha(ws: Worksheet, row: int, fecha: date) -> None:
    cell = ws.cell(row, 1)
    cell.value = format_fecha_es(fecha)
    cell.number_format = "@"


def _write_formulas(ws: Worksheet, row: int, fee_rate_cell: str) -> None:
    ws.cell(row, 5).value = f"=D{row}*{fee_rate_cell}"  # FEE
    ws.cell(row, 6).value = f"=D{row}-E{row}"  # TARJETA NETA
    ws.cell(row, 10).value = f"=C{row}+F{row}+G{row}-I{row}"  # TOTAL
    ws.cell(row, 13).value = f"=J{row}-L{row}"  # VENTAS SIN ITBIS


def _write_stub(ws: Worksheet, row: int, fecha: date, ubicacion: str, fee_rate_cell: str) -> None:
    """Placeholder row for a location on a date (amounts 0 until filled)."""
    _write_fecha(ws, row, fecha)
    ws.cell(row, 2).value = ubicacion
    ws.cell(row, 3).value = 0  # EFECTIVO
    ws.cell(row, 4).value = 0  # TARJETA
    ws.cell(row, 7).value = 0  # TRANSFERENCIAS
    ws.cell(row, 8).value = None  # PEDIDOS YA
    ws.cell(row, 9).value = 0  # NOTAS CREDITO
    ws.cell(row, 12).value = 0  # ITBS
    _write_formulas(ws, row, fee_rate_cell)
    _apply_row_formats(ws, row)


def apply_row(ws: Worksheet, row: int, data: DailySaleRow, fee_rate_cell: str) -> None:
    _write_fecha(ws, row, data.fecha)
    ws.cell(row, 2).value = data.ubicacion
    ws.cell(row, 3).value = data.efectivo
    ws.cell(row, 4).value = data.tarjeta_bruta
    ws.cell(row, 7).value = data.transferencias
    if data.pedidos_ya is not None:
        ws.cell(row, 8).value = data.pedidos_ya
    ws.cell(row, 9).value = data.notas_credito
    if data.itbs_cobrado is not None:
        ws.cell(row, 12).value = data.itbs_cobrado
    elif ws.cell(row, 12).value is None:
        ws.cell(row, 12).value = 0
    # Never overwrite deposito (N / col 14) if already filled
    _write_formulas(ws, row, fee_rate_cell)
    _apply_row_formats(ws, row)


def _ensure_day_block(
    ws: Worksheet,
    fecha: date,
    ubicaciones: list[str],
    fee_rate_cell: str,
) -> dict[str, int]:
    """Ensure one row per ubicacion for fecha, in order, chronologically placed."""
    existing = _rows_for_fecha(ws, fecha)
    if len(existing) == len(ubicaciones) and all(u in existing for u in ubicaciones):
        return existing

    if not existing:
        insert_at = _chronological_insert_row(ws, fecha)
        ws.insert_rows(insert_at, amount=len(ubicaciones))
        row_map: dict[str, int] = {}
        for offset, ubic in enumerate(ubicaciones):
            row = insert_at + offset
            _write_stub(ws, row, fecha, ubic, fee_rate_cell)
            row_map[ubic] = row
        return row_map

    # Partial day exists — add missing ubicaciones after the day's last row,
    # preferring configured order when possible.
    for ubic in ubicaciones:
        if ubic in existing:
            continue
        # Insert after the last row currently belonging to this fecha
        current = _rows_for_fecha(ws, fecha)
        after = max(current.values()) + 1
        ws.insert_rows(after, amount=1)
        _write_stub(ws, after, fecha, ubic, fee_rate_cell)

    return _rows_for_fecha(ws, fecha)


def upsert_daily_sale(
    excel_path: str | Path,
    sheet_name: str,
    data: DailySaleRow,
    fee_rate_cell: str = "Setup!$B$6",
    ubicaciones: list[str] | None = None,
) -> dict[str, Any]:
    path = Path(excel_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")
    if not path.is_file():
        raise ValueError(f"Excel path is not a file: {path}")

    wb = load_workbook(path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name} in {path}")
    ws = wb[sheet_name]
    locations = ubicaciones or DEFAULT_UBICACIONES
    if data.ubicacion not in locations:
        # Still allow the PDF ubicacion; put it first if unknown
        locations = [data.ubicacion, *[u for u in locations if u != data.ubicacion]]

    before = _rows_for_fecha(ws, data.fecha)
    row_map = _ensure_day_block(ws, data.fecha, locations, fee_rate_cell)
    row = row_map[data.ubicacion]
    action = "updated" if data.ubicacion in before else "inserted"
    apply_row(ws, row, data, fee_rate_cell)
    # Keep sibling stub rows on the same date visually consistent
    for ubic, sibling_row in row_map.items():
        if ubic != data.ubicacion:
            _write_fecha(ws, sibling_row, data.fecha)
            _apply_row_formats(ws, sibling_row)
            _write_formulas(ws, sibling_row, fee_rate_cell)

    # Save in place. Do NOT delete/replace the OneDrive file (that triggers
    # "Deleted Files Are Removed Everywhere" and can drop the workbook).
    wb.save(path)
    wb.close()

    return {
        "action": action,
        "row": row,
        "fecha": data.fecha.isoformat(),
        "fecha_display": format_fecha_es(data.fecha),
        "ubicacion": data.ubicacion,
        "excel_path": str(path),
        "day_rows": row_map,
        "created_locations": [u for u in locations if u not in before],
    }


def upsert_with_retry(
    excel_path: str | Path,
    sheet_name: str,
    data: DailySaleRow,
    fee_rate_cell: str = "Setup!$B$6",
    retries: int = 5,
    retry_seconds: float = 2.0,
    pending_path: str | Path | None = None,
    ubicaciones: list[str] | None = None,
) -> dict[str, Any]:
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            result = upsert_daily_sale(
                excel_path,
                sheet_name,
                data,
                fee_rate_cell=fee_rate_cell,
                ubicaciones=ubicaciones,
            )
            result["attempts"] = attempt
            return result
        except PermissionError as exc:
            last_error = exc
            time.sleep(retry_seconds)
        except OSError as exc:
            last_error = exc
            time.sleep(retry_seconds)

    if pending_path is not None:
        queue_pending(pending_path, data)
        return {
            "action": "queued",
            "error": str(last_error),
            "fecha": data.fecha.isoformat(),
            "ubicacion": data.ubicacion,
        }
    raise RuntimeError(f"Failed to update Excel after {retries} attempts: {last_error}")


def queue_pending(pending_path: str | Path, data: DailySaleRow) -> None:
    path = Path(pending_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    items: list[dict[str, Any]] = []
    if path.exists():
        items = json.loads(path.read_text(encoding="utf-8"))
    payload = asdict(data)
    payload["fecha"] = data.fecha.isoformat()
    items.append(payload)
    path.write_text(json.dumps(items, indent=2), encoding="utf-8")


def flush_pending(
    pending_path: str | Path,
    excel_path: str | Path,
    sheet_name: str,
    fee_rate_cell: str = "Setup!$B$6",
    retries: int = 5,
    retry_seconds: float = 2.0,
    ubicaciones: list[str] | None = None,
) -> list[dict[str, Any]]:
    path = Path(pending_path)
    if not path.exists():
        return []
    items = json.loads(path.read_text(encoding="utf-8"))
    remaining: list[dict[str, Any]] = []
    results: list[dict[str, Any]] = []
    for item in items:
        data = DailySaleRow(
            fecha=date.fromisoformat(item["fecha"]),
            ubicacion=item["ubicacion"],
            efectivo=item["efectivo"],
            tarjeta_bruta=item["tarjeta_bruta"],
            transferencias=item["transferencias"],
            notas_credito=item["notas_credito"],
            itbs_cobrado=item.get("itbs_cobrado"),
            pedidos_ya=item.get("pedidos_ya"),
        )
        result = upsert_with_retry(
            excel_path,
            sheet_name,
            data,
            fee_rate_cell=fee_rate_cell,
            retries=retries,
            retry_seconds=retry_seconds,
            pending_path=None,
            ubicaciones=ubicaciones,
        )
        if result.get("action") == "queued":
            remaining.append(item)
        else:
            results.append(result)
    if remaining:
        path.write_text(json.dumps(remaining, indent=2), encoding="utf-8")
    else:
        path.unlink(missing_ok=True)
    return results
