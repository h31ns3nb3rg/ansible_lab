"""Parse AdControl 'Informe avanzado de cierres de caja' Excel exports."""

from __future__ import annotations

import logging
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
    "fondo_caja": ("fondo de caja",),
    # Sales cash for Ventas Diarias EFECTIVO (excludes fondo de caja float)
    "pago_efectivo": ("pago en efectivo", "pago en efectivos"),
    "hora_apertura": ("hora de apertura",),
    "tarjeta": ("total en pago con tarjeta", "pago con tarjeta"),
    "transferencia": (
        "transeferencia bancaria",  # AdControl typo
        "transferencia bancaria",
    ),
    "pedidos_ya": ("total en pedidosya", "pedidosya", "pedidos ya"),
    "otros_pagos": ("total en otros pagos", "otros pagos"),
}

# Expected float per store when export omits Fondo de Caja (fallback only).
_DEFAULT_FONDO_BY_UBICACION: dict[str, float] = {
    "La Cata SJM": 5000.0,
    "La Cata DF": 2500.0,
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
        "hora_apertura",
        "tarjeta",
        "transferencia",
        "otros_pagos",
    )
    missing = [k for k in required if k not in resolved]
    # Need either Pago en efectivo, or Cantidad de cierre (+ optional Fondo) to derive it.
    if "pago_efectivo" not in resolved and "cantidad_cierre" not in resolved:
        missing.append("pago_efectivo|cantidad_cierre")
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
    fondo_caja: float = 0.0
    cantidad_cierre: float | None = None
    usuario: str | None = None
    estado: str | None = None


def _efectivo_from_row(
    row: tuple[Any, ...],
    cols: dict[str, int],
    ubicacion: str,
) -> tuple[float, float, float | None]:
    """Return (pago_en_efectivo, fondo_caja, cantidad_cierre).

    Ventas Diarias EFECTIVO uses AdControl **Pago en efectivo** (sales cash).
    That equals Cantidad de cierre − Fondo de Caja (float is not sales).
    """
    cantidad: float | None = None
    if "cantidad_cierre" in cols:
        cantidad = _parse_money(row[cols["cantidad_cierre"]])

    if "fondo_caja" in cols:
        fondo = _parse_money(row[cols["fondo_caja"]])
    else:
        fondo = _DEFAULT_FONDO_BY_UBICACION.get(ubicacion, 0.0)

    if "pago_efectivo" in cols:
        efectivo = _parse_money(row[cols["pago_efectivo"]])
    elif cantidad is not None:
        efectivo = max(0.0, cantidad - fondo)
    else:
        raise ValueError("Cierres row missing Pago en efectivo and Cantidad de cierre")

    if cantidad is not None and "pago_efectivo" in cols:
        expected = max(0.0, round(cantidad - fondo, 2))
        if abs(expected - round(efectivo, 2)) > 0.01:
            logging.warning(
                "Cierres mismatch %s: cantidad_cierre=%s fondo=%s "
                "pago_efectivo=%s (expected cantidad-fondo=%s)",
                ubicacion,
                cantidad,
                fondo,
                efectivo,
                expected,
            )
    return efectivo, fondo, cantidad


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
            ubicacion = _map_ubicacion(ubic_raw)
            efectivo, fondo, cantidad = _efectivo_from_row(row, cols, ubicacion)
            shifts.append(
                CierreShift(
                    fecha=fecha,
                    ubicacion=ubicacion,
                    # Ventas Diarias EFECTIVO = Pago en efectivo (not Cantidad de cierre)
                    efectivo=efectivo,
                    tarjeta=_parse_money(row[cols["tarjeta"]]),
                    transferencias=transferencia + otros,
                    pedidos_ya=pedidos,
                    fondo_caja=fondo,
                    cantidad_cierre=cantidad,
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


def is_cierres_workbook(path: str | Path) -> bool:
    """True if workbook looks like AdControl 'Informe avanzado de cierres de caja'."""
    path = Path(path)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    # Ignore Excel lock / temp files
    if path.name.startswith("~$") or path.name.startswith("."):
        return False
    # Fast path: common export filename (spaces, -2 copies, etc. are fine)
    name = path.name.lower()
    if ("cierres" in name and "caja" in name) or (
        "informe" in name and "cierres" in name
    ):
        return True
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.active
            header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if not header_row:
                return False
            _resolve_headers(list(header_row))
            return True
        finally:
            wb.close()
    except Exception as exc:  # noqa: BLE001
        logging.debug("Not a cierres workbook %s: %s", path.name, exc)
        return False


def resolve_cierres_path(path: str | Path, base_dir: str | Path | None = None) -> Path:
    """Resolve a cierres Excel path; try inbox/ when given a bare filename."""
    raw = Path(path).expanduser()
    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.append(raw)
    else:
        candidates.append(Path.cwd() / raw)
        if base_dir is not None:
            base = Path(base_dir)
            candidates.append(base / raw)
            candidates.append(base / "inbox" / raw.name)
            # Also allow just the filename while cwd is elsewhere
            candidates.append(base / "inbox" / Path(path).name)

    seen: set[Path] = set()
    for cand in candidates:
        try:
            resolved = cand.resolve()
        except OSError:
            continue
        if resolved in seen:
            continue
        seen.add(resolved)
        if resolved.is_file():
            return resolved

    searched = ", ".join(str(c) for c in candidates)
    raise FileNotFoundError(
        f"Cierres Excel not found: {path!r}. Looked in: {searched}. "
        "Put the file in inbox/ or pass a full path in quotes "
        '(spaces in the name require quotes).'
    )
