"""Parse VENTAS Y COBROS PDFs for Totales Generales row."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass
class TotalesGenerales:
    fecha: datetime
    efectivo: float
    tarjeta: float
    cheque: float
    otros_pagos: float
    nota_credito: float
    valor_neto: float
    pendiente: float
    itbis: float | None = None
    cantidad: int | None = None


_AMOUNT_RE = re.compile(r"-?[\d,]+\.\d{2}")
_DATE_RE = re.compile(r"\b(\d{2}/\d{2}/\d{4})\b")
_TM_TJ_RE = re.compile(
    rb"([0-9.]+) ([0-9.]+) Tm \(((?:\\.|[^\\)])*)\) Tj"
)


def _unescape_pdf_text(raw: bytes) -> str:
    text = raw.decode("latin-1")
    return text.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")


def _parse_amount(text: str) -> float:
    cleaned = text.strip().replace(",", "")
    return float(cleaned)


def _extract_positioned_text(pdf_bytes: bytes) -> list[tuple[float, float, str]]:
    streams = re.findall(rb"stream\r?\n(.*?)\r?\nendstream", pdf_bytes, re.S)
    if not streams:
        raise ValueError("No PDF content streams found")

    items: list[tuple[float, float, str]] = []
    for stream in streams:
        for match in _TM_TJ_RE.finditer(stream):
            x = float(match.group(1))
            y = float(match.group(2))
            text = _unescape_pdf_text(match.group(3)).strip()
            if text:
                items.append((y, x, text))
    if not items:
        raise ValueError("No text operators found in PDF")
    return items


def _find_fecha(items: list[tuple[float, float, str]]) -> datetime:
    # Prefer dates near the header (high Y), not cajero lines only.
    candidates: list[tuple[float, datetime]] = []
    for y, _x, text in items:
        for match in _DATE_RE.finditer(text):
            candidates.append((y, datetime.strptime(match.group(1), "%d/%m/%Y")))
    if not candidates:
        raise ValueError("Could not find report date (dd/mm/yyyy) in PDF")
    candidates.sort(key=lambda pair: pair[0], reverse=True)
    return candidates[0][1]


def _amounts_on_line(
    items: list[tuple[float, float, str]], y_target: float, tol: float = 1.5
) -> list[tuple[float, float]]:
    row = [(x, _parse_amount(text)) for y, x, text in items if abs(y - y_target) <= tol and _AMOUNT_RE.fullmatch(text.replace(" ", ""))]
    # Also accept amounts with leading spaces already stripped
    if not row:
        row = []
        for y, x, text in items:
            if abs(y - y_target) > tol:
                continue
            compact = text.replace(" ", "")
            if _AMOUNT_RE.fullmatch(compact):
                row.append((x, _parse_amount(compact)))
    row.sort(key=lambda pair: pair[0])
    return row


def parse_totales_generales(pdf_path: str | Path) -> TotalesGenerales:
    """Extract Totales Generales (+ optional Itbis) from a ventas PDF."""
    path = Path(pdf_path)
    pdf_bytes = path.read_bytes()
    items = _extract_positioned_text(pdf_bytes)
    fecha = _find_fecha(items)

    totales_y = None
    for y, _x, text in items:
        if "Totales Generales" in text:
            totales_y = y
            break
    if totales_y is None:
        raise ValueError("Could not find 'Totales Generales' in PDF")

    amounts = _amounts_on_line(items, totales_y)
    # Expected order aligned to headers after label:
    # Cr Factura, Efectivo, Tarjeta, Cheque, Otros Pagos, Nota Cr., Valor Neto, Pendiente
    if len(amounts) < 8:
        raise ValueError(
            f"Expected at least 8 amounts on Totales Generales line, found {len(amounts)}: {amounts}"
        )

    _cr, efectivo, tarjeta, cheque, otros, nota, valor_neto, pendiente = [
        value for _x, value in amounts[:8]
    ]

    itbis = None
    cantidad = None
    for _y, _x, text in items:
        if text.lower().startswith("itbis"):
            # value may be on same item or nearby
            pass
    for y, x, text in items:
        if "Itbis" in text or text.lower().startswith("itbis"):
            # look for amount near this y
            nearby = _amounts_on_line(items, y, tol=2.0)
            if nearby:
                itbis = nearby[0][1]
            else:
                # sometimes "Itbis --> 8,521.03" split; check next tokens on line
                for yy, xx, tt in items:
                    if abs(yy - y) <= 2 and xx > x and _AMOUNT_RE.fullmatch(tt.replace(" ", "")):
                        itbis = _parse_amount(tt.replace(" ", ""))
                        break
        if text.lower().startswith("cantidad"):
            m = re.search(r"(\d+)", text)
            if m:
                cantidad = int(m.group(1))

    # Dedicated pass for Itbis amount if label and value are separate
    if itbis is None:
        for y, x, text in items:
            if text.strip() in {"Itbis -->", "Itbis-->", "Itbis"} or text.startswith("Itbis"):
                for yy, xx, tt in sorted(items, key=lambda t: t[1]):
                    if abs(yy - y) <= 2.0 and xx >= x:
                        compact = tt.replace(" ", "")
                        if _AMOUNT_RE.fullmatch(compact):
                            itbis = _parse_amount(compact)
                            break

    return TotalesGenerales(
        fecha=fecha,
        efectivo=efectivo,
        tarjeta=tarjeta,
        cheque=cheque,
        otros_pagos=otros,
        nota_credito=nota,
        valor_neto=valor_neto,
        pendiente=pendiente,
        itbis=itbis,
        cantidad=cantidad,
    )
