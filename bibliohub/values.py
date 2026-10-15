"""Testo pulito per valori letti da Excel.

Pandas legge una colonna di interi con celle vuote come float: 15 diventa 15.0
e, una volta convertito in stringa, resta «15.0» in Excel, Word e PDF.
"""

import re
from datetime import date, datetime

_TRAILING_ZERO = re.compile(r"-?\d+\.0+")


def as_text(value) -> str:
    """Rende il valore come testo, senza il «.0» degli interi letti come float."""
    if value is None or isinstance(value, bool):
        return "" if value is None else str(value)
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, str):
        text = value.strip()
    elif isinstance(value, int):
        return str(value)
    elif isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            return str(int(value))
        text = format(value, "f").rstrip("0").rstrip(".")
        return text
    else:
        text = str(value).strip()
        if text.lower() in {"nan", "nat", "none", "<na>"}:
            return ""
    if _TRAILING_ZERO.fullmatch(text):
        return text.split(".", 1)[0]
    return text
