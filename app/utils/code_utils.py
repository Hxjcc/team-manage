from __future__ import annotations

import re
from typing import Any, Optional

_STANDARD_CODE_RE = re.compile(r"[A-Za-z0-9]{4}(?:-[A-Za-z0-9]{4}){3}")
_GENERIC_CODE_RE = re.compile(r"(?i)(?=.*[A-Z])[A-Z0-9-]{8,32}")


def normalize_code_input(value: Any) -> Optional[str]:
    """
    Normalize code-like input.

    Accepts inputs like:
    - "ABCD-EFGH-IJKL-MNOP"
    - "ABCD-EFGH-IJKL-MNOP\\n￥12.5"
    - "ABCD-EFGH-IJKL-MNOP 已支付"
    """
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return ""

    match = _STANDARD_CODE_RE.search(text)
    if match:
        return match.group(0).strip()

    match = _GENERIC_CODE_RE.search(text)
    if match:
        return match.group(0).strip()

    for line in re.split(r"\r?\n", text):
        line = line.strip()
        if not line:
            continue
        return line.split()[0]

    return text

