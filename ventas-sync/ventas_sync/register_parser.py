"""Parse AdControl register-report PDFs for DF / SJM stores."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from pypdf import PdfReader

from .excel_updater import DailySaleRow

_MONTHS = {
    "ene": 1,
    "enero": 1,
    "feb": 2,
    "febrero": 2,
    "mar": 3,
    "marzo": 3,
    "abr": 4,
    "abril": 4,
    "may": 5,
    "mayo": 5,
    "jun": 6,
    "junio": 6,
    "jul": 7,
    "julio": 7,
    "ago": 8,
    "agosto": 8,
    "sep": 9,
    "sept": 9,
    "septiembre": 9,
    "set": 9,
    "setiembre": 9,
    "oct": 10,
    "octubre": 10,
    "nov": 11,
    "noviembre": 11,
    "dic": 12,
    "diciembre": 12,
}

_LOCATION_MAP = {
    "SAN JUAN": "La Cata SJM",
    "DEFILLO": "La Cata DF",
}


def _strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn"
    )


def _parse_rd_amount(text: str) -> float | None:
    match = re.search(r"RD\$\s*([-0-9,]+\.\d{2})", text)
    if not match:
        return None
    return float(match.group(1).replace(",", ""))


def _extract_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    chunks: list[str] = []
    for page in reader.pages:
        chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def _find_labeled_amount(text: str, label: str) -> float | None:
    """Find first RD$ amount on the same line / nearby after a label."""
    # Label may wrap; allow limited gap before amount
    pattern = re.compile(
        re.escape(label) + r".{0,80}?RD\$\s*([-0-9,]+\.\d{2})",
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(text)
    if not match:
        return None
    return float(match.group(1).replace(",", ""))


def _parse_caja_date(text: str) -> date:
    # Detalles de caja ( 26 de jul. 2026 9:16 AM - 26 de jul. 2026 4:39 PM )
    match = re.search(
        r"Detalles de caja\s*\(\s*(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóú\.]+)\s+(\d{4})",
        text,
        re.IGNORECASE,
    )
    if not match:
        # Fallback: 26-07-2026 style if present
        m2 = re.search(r"\b(\d{2})-(\d{2})-(\d{4})\b", text)
        if m2:
            return date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
        raise ValueError("Could not find cash-register date in PDF")

    day = int(match.group(1))
    month_raw = _strip_accents(match.group(2)).lower().strip(".")
    year = int(match.group(3))
    month = _MONTHS.get(month_raw)
    if not month:
        raise ValueError(f"Unknown month in register PDF: {match.group(2)!r}")
    return date(year, month, day)


def _parse_ubicacion(text: str) -> str:
    # Ubicación comercial: ... (SAN JUAN) or (DEFILLO)
    compact = _strip_accents(text).upper()
    if "SAN JUAN" in compact:
        return "La Cata SJM"
    if "DEFILLO" in compact:
        return "La Cata DF"
    raise ValueError(
        "Could not map location in register PDF. Expected SAN JUAN or DEFILLO in "
        "'Ubicación comercial'."
    )


@dataclass
class RegisterReport:
    fecha: date
    ubicacion: str
    efectivo: float
    tarjeta: float
    transferencias: float
    pedidos_ya: float
    otros_pagos: float
    itbis: float | None
    usuario: str | None = None
    source: str = "adcontrol_register"


def parse_register_report(pdf_path: str | Path) -> RegisterReport:
    path = Path(pdf_path)
    text = _extract_text(path)
    if "Detalles de caja" not in text and "register-report" not in text.lower():
        raise ValueError("Not an AdControl register-report PDF")

    fecha = _parse_caja_date(text)
    ubicacion = _parse_ubicacion(text)

    efectivo = _find_labeled_amount(text, "Efectivo del Dia")
    if efectivo is None:
        efectivo = _find_labeled_amount(text, "Efectivo Neto")
    if efectivo is None:
        raise ValueError("Could not find 'Efectivo del Dia' / 'Efectivo Neto'")

    tarjeta = _find_labeled_amount(text, "Pago con tarjeta")
    if tarjeta is None:
        tarjeta = 0.0

    transferencias = _find_labeled_amount(text, "Transferencia bancaria") or 0.0
    otros = _find_labeled_amount(text, "Otros pagos") or 0.0
    pedidos = _find_labeled_amount(text, "PedidosYA") or 0.0
    # Also accept "Pedidos YA"
    if pedidos == 0.0:
        pedidos = _find_labeled_amount(text, "Pedidos YA") or 0.0

    itbis = _find_labeled_amount(text, "Impuesto:")
    if itbis is None:
        itbis = _find_labeled_amount(text, "Impuesto")

    usuario = None
    um = re.search(r"Usuario:\s*(.+)", text)
    if um:
        usuario = um.group(1).strip().split("\n")[0].strip()

    return RegisterReport(
        fecha=fecha,
        ubicacion=ubicacion,
        efectivo=efectivo,
        tarjeta=tarjeta,
        transferencias=transferencias + otros,
        pedidos_ya=pedidos,
        otros_pagos=otros,
        itbis=itbis,
        usuario=usuario,
    )


def register_to_row(report: RegisterReport) -> DailySaleRow:
    return DailySaleRow(
        fecha=report.fecha,
        ubicacion=report.ubicacion,
        efectivo=report.efectivo,
        tarjeta_bruta=report.tarjeta,
        transferencias=report.transferencias,
        notas_credito=0.0,
        itbs_cobrado=report.itbis,
        pedidos_ya=report.pedidos_ya,
    )


def detect_pdf_kind(pdf_path: str | Path) -> str:
    """Return 'lmf_ventas', 'adcontrol_register', or 'unknown'."""
    path = Path(pdf_path)
    # Fast path: LMF text PDFs have plain operators
    raw = path.read_bytes()
    if b"Totales Generales" in raw or b"VENTAS Y COBROS" in raw:
        return "lmf_ventas"
    try:
        text = _extract_text(path)
    except Exception:
        return "unknown"
    if "Detalles de caja" in text or "Ubicación comercial" in text or "Ubicacion comercial" in _strip_accents(text):
        return "adcontrol_register"
    if "Totales Generales" in text or "VENTAS Y COBROS" in text:
        return "lmf_ventas"
    return "unknown"
