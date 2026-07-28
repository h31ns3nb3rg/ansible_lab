"""Parse AdControl 'Informe avanzado de cierres de caja' Excel exports."""

from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .excel_updater import DailySaleRow

_LOCATION_MARKERS = (
    ("SAN JUAN", "La Cata SJM"),
    ("DEFILLO", "La Cata DF"),
)

# Flexible header aliases (export typos / renames).
_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "ubicacion": ("ubicacion", "ubicacion comercial"),
    "estado": ("estado",),
    "cantidad_cierre": ("cantidad de cierre",),
    "hora_apertura": ("hora de apertura",),
    "tarjeta": ("total en pago con tarjeta", "pago con tarjeta"),
    "transferencia": (
        "transeferencia bancaria",  # AdControl typo
        "transferencia bancaria",
    ),
    "pedidos_ya": ("total en pedidosya", "pedidosya", "pedidos ya"),
    "otros_pagos": ("total en otros pagos", "otros pagos"),
}


def _strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn"
    )


def _norm_header(value: Any) -> str:
    if value is None:
        return ""
    text = _strip_accents(str(value)).lower().strip()
    return " ".join(text.split())


def _parse_money(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    # Prefer last money-looking token (handles "RD$ 33,597.60")
    matches = re.findall(r"[-0-9,]+\.\d{2}", text)
    if matches:
        return float(matches[-1].replace(",", ""))
    cleaned = (
        text.replace("RD$", "")
        .replace("rd$", "")
        .replace(",", "")
        .replace(" ", "")
        .strip()
    )
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError as exc:
        raise ValueError(f"Could not parse money value: {value!r}") from exc


def _parse_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)):
        from openpyxl.utils.datetime import from_excel

        return from_excel(value)
    text = str(value).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ValueError(f"Could not parse datetime: {value!r}")


def _map_ubicacion(raw: Any) -> str:
    compact = _strip_accents(str(raw or "")).upper()
    has_sjm = bool(re.search(r"\(\s*SAN\s+JUAN\s*\)", compact)) or (
        "SAN JUAN" in compact and "DEFILLO" not in compact
    )
    has_df = bool(re.search(r"\(\s*DEFILLO\s*\)", compact)) or (
        "DEFILLO" in compact and "SAN JUAN" not in compact
    )
    if has_sjm and has_df:
        raise ValueError(f"Ambiguous ubicación in cierres export: {raw!r}")
    if has_sjm:
        return "La Cata SJM"
    if has_df:
        return "La Cata DF"
    raise ValueError(f"Unknown ubicación in cierres export: {raw!r}")


def _resolve_headers(headers: list[Any]) -> dict[str, int]:
    normed = [_norm_header(h) for h in headers]
    resolved: dict[str, int] = {}
    for key, aliases in _HEADER_ALIASES.items():
        for idx, name in enumerate(normed):
            if name in aliases:
                resolved[key] = idx
                break
    required = (
        "ubicacion",
        "cantidad_cierre",
        "hora_apertura",
        "tarjeta",
        "transferencia",
        "otros_pagos",
    )
    missing = [k for k in required if k not in resolved]
    if missing:
        raise ValueError(
            "Cierres Excel missing required columns: "
            + ", ".join(missing)
            + f". Found headers: {headers!r}"
        )
    return resolved


@dataclass
class CierreShift:
    fecha: date
    ubicacion: str
    efectivo: float
    tarjeta: float
    transferencias: float
    pedidos_ya: float
    usuario: str | None = None
    estado: str | None = None


def parse_cierres_workbook(path: str | Path) -> list[CierreShift]:
    """Parse each closed cash-register shift from the advanced cierres Excel."""
    wb = load_workbook(Path(path), data_only=True, read_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration as exc:
            raise ValueError("Cierres Excel is empty") from exc
        cols = _resolve_headers(list(header_row))
        shifts: list[CierreShift] = []
        for row in rows:
            if row is None or all(v is None or str(v).strip() == "" for v in row):
                continue
            ubic_raw = row[cols["ubicacion"]]
            if ubic_raw is None or str(ubic_raw).strip() == "":
                continue
            estado = row[cols["estado"]] if "estado" in cols else None
            if estado is not None and str(estado).strip().lower() not in {"", "close", "cerrado"}:
                # Skip open / incomplete registers
                continue
            fecha = _parse_datetime(row[cols["hora_apertura"]]).date()
            transferencia = _parse_money(row[cols["transferencia"]])
            otros = _parse_money(row[cols["otros_pagos"]])
            pedidos = (
                _parse_money(row[cols["pedidos_ya"]]) if "pedidos_ya" in cols else 0.0
            )
            usuario = None
            # Usuario is usually column 1; resolve if present among headers
            for idx, h in enumerate(header_row):
                if _norm_header(h) == "usuario":
                    raw_user = row[idx]
                    if raw_user is not None:
                        usuario = str(raw_user).strip() or None
                    break
            shifts.append(
                CierreShift(
                    fecha=fecha,
                    ubicacion=_map_ubicacion(ubic_raw),
                    # Matches register PDF "Efectivo del Dia" / Cantidad de cierre
                    efectivo=_parse_money(row[cols["cantidad_cierre"]]),
                    tarjeta=_parse_money(row[cols["tarjeta"]]),
                    transferencias=transferencia + otros,
                    pedidos_ya=pedidos,
                    usuario=usuario,
                    estado=str(estado).strip() if estado is not None else None,
                )
            )
        return shifts
    finally:
        wb.close()


def aggregate_cierres(shifts: list[CierreShift]) -> list[DailySaleRow]:
    """Sum shifts by (fecha, ubicacion) into Ventas Diarias rows."""
    buckets: dict[tuple[date, str], DailySaleRow] = {}
    for shift in shifts:
        key = (shift.fecha, shift.ubicacion)
        if key not in buckets:
            buckets[key] = DailySaleRow(
                fecha=shift.fecha,
                ubicacion=shift.ubicacion,
                efectivo=shift.efectivo,
                tarjeta_bruta=shift.tarjeta,
                transferencias=shift.transferencias,
                notas_credito=0.0,
                itbs_cobrado=None,  # not present in cierres export
                pedidos_ya=shift.pedidos_ya,
            )
            continue
        existing = buckets[key]
        buckets[key] = DailySaleRow(
            fecha=existing.fecha,
            ubicacion=existing.ubicacion,
            efectivo=existing.efectivo + shift.efectivo,
            tarjeta_bruta=existing.tarjeta_bruta + shift.tarjeta,
            transferencias=existing.transferencias + shift.transferencias,
            notas_credito=existing.notas_credito,
            itbs_cobrado=existing.itbs_cobrado,
            pedidos_ya=(existing.pedidos_ya or 0.0) + shift.pedidos_ya,
        )
    return sorted(buckets.values(), key=lambda r: (r.fecha, r.ubicacion))


def summarize_cierres(rows: list[DailySaleRow]) -> dict[str, Any]:
    by_ubic: dict[str, int] = defaultdict(int)
    for row in rows:
        by_ubic[row.ubicacion] += 1
    return {
        "days_locations": len(rows),
        "by_ubicacion": dict(by_ubic),
        "date_from": rows[0].fecha.isoformat() if rows else None,
        "date_to": rows[-1].fecha.isoformat() if rows else None,
    }
