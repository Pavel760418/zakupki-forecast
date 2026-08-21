# -*- coding: utf-8 -*-
"""Однократная генерация config/quantum_library.py из файла привязки SKU."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from data.quantum_parse import parse_quantum_from_name  # noqa: E402
from utils.helpers import normalize_text, safe_str  # noqa: E402

SRC = Path(r"C:\Users\Администратор\Documents\МР\новый период работы\Закупка\Привязка_SKU_к_контрагенту.xlsx")
OUT = ROOT / "config" / "quantum_library.py"


def _load_items(path: Path) -> list[dict]:
    raw = pd.read_excel(path, header=None)
    # структура как в supplier_mapping: номенклатура col0, артикул 7, код 8, штрихкод 9
    items = []
    for i in range(8, len(raw)):
        name = safe_str(raw.iloc[i, 0]).strip()
        if not name or "итог" in name.casefold():
            continue
        article = safe_str(raw.iloc[i, 7]).strip() if raw.shape[1] > 7 else ""
        code = safe_str(raw.iloc[i, 8]).strip() if raw.shape[1] > 8 else ""
        barcode = safe_str(raw.iloc[i, 9]).strip() if raw.shape[1] > 9 else ""
        # пропускаем строки-группировки без кода/артикула/штрихкода (группы номенклатуры)
        if not (article or code or barcode):
            # но оставляем, если в имени явно есть квант и похоже на товар (есть цифры)
            if not re.search(r"\d", name):
                continue
        sku_key = normalize_text(article) or normalize_text(code) or normalize_text(barcode) or normalize_text(name)
        q = parse_quantum_from_name(name)
        if not q:
            continue
        items.append(
            {
                "sku_key": sku_key,
                "name": name,
                "name_key": normalize_text(name),
                "article": article,
                "code": code,
                "barcode": barcode,
                "quantum": int(q),
            }
        )
    return items


def main() -> None:
    items = _load_items(SRC)
    by_sku: dict[str, int] = {}
    by_name: dict[str, int] = {}
    samples: list[tuple[int, str]] = []
    for it in items:
        q = it["quantum"]
        if it["sku_key"] and it["sku_key"] not in by_sku:
            by_sku[it["sku_key"]] = q
        if it["name_key"] and it["name_key"] not in by_name:
            by_name[it["name_key"]] = q
        if len(samples) < 40:
            samples.append((q, it["name"]))

    lines = [
        '"""',
        "Библиотека квантов (шт в упаковке/коробе) — зафиксирована в коде.",
        "",
        "Источник первичного наполнения: Привязка_SKU_к_контрагенту.xlsx",
        "(квант извлекается из наименования: *N или /N, 1/N).",
        "",
        "Как править:",
        "1) QUANTUM_OVERRIDES_BY_SKU_KEY / QUANTUM_OVERRIDES_BY_NAME_KEY — точечные правки",
        "   (имеют приоритет над авто-словарём);",
        "2) либо обновить словари QUANTUM_BY_* и перезапустить приложение.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "from typing import Dict, Optional",
        "",
        "from data.quantum_parse import ceil_to_quantum, parse_quantum_from_name",
        "from utils.helpers import normalize_text, safe_str",
        "",
        "# Ручные правки (приоритет над автозаполнением).",
        "QUANTUM_OVERRIDES_BY_SKU_KEY: Dict[str, int] = {",
        "    # \"00-00001234\": 24,",
        "}",
        "",
        "QUANTUM_OVERRIDES_BY_NAME_KEY: Dict[str, int] = {",
        "    # normalize_text(\"пример названия\"): 12,",
        "}",
        "",
        f"# Авто из файла привязки: {len(by_sku)} ключей SKU, {len(by_name)} имён.",
        "QUANTUM_BY_SKU_KEY: Dict[str, int] = {",
    ]
    for k in sorted(by_sku.keys()):
        lines.append(f"    {json.dumps(k, ensure_ascii=False)}: {by_sku[k]},")
    lines.append("}")
    lines.append("")
    lines.append("QUANTUM_BY_NAME_KEY: Dict[str, int] = {")
    for k in sorted(by_name.keys()):
        lines.append(f"    {json.dumps(k, ensure_ascii=False)}: {by_name[k]},")
    lines.append("}")
    lines.append("")
    lines.extend(
        [
            "",
            "def lookup_quantum(",
            "    *,",
            "    sku_key: object = \"\",",
            "    name: object = \"\",",
            "    sku: object = \"\",",
            "    article: object = \"\",",
            "    code: object = \"\",",
            "    barcode: object = \"\",",
            ") -> int:",
            "    \"\"\"Квант для позиции: override → словарь → разбор имени → 1.\"\"\"",
            "    candidates = [",
            "        normalize_text(sku_key),",
            "        normalize_text(sku),",
            "        normalize_text(article),",
            "        normalize_text(code),",
            "        normalize_text(barcode),",
            "    ]",
            "    for key in candidates:",
            "        if key and key in QUANTUM_OVERRIDES_BY_SKU_KEY:",
            "            return int(QUANTUM_OVERRIDES_BY_SKU_KEY[key])",
            "    name_key = normalize_text(name)",
            "    if name_key and name_key in QUANTUM_OVERRIDES_BY_NAME_KEY:",
            "        return int(QUANTUM_OVERRIDES_BY_NAME_KEY[name_key])",
            "    for key in candidates:",
            "        if key and key in QUANTUM_BY_SKU_KEY:",
            "            return int(QUANTUM_BY_SKU_KEY[key])",
            "    if name_key and name_key in QUANTUM_BY_NAME_KEY:",
            "        return int(QUANTUM_BY_NAME_KEY[name_key])",
            "    parsed = parse_quantum_from_name(name)",
            "    if parsed:",
            "        return int(parsed)",
            "    return 1",
            "",
            "",
            "def apply_quantum_ceil(qty: float, quantum: int | None) -> float:",
            "    return ceil_to_quantum(qty, quantum)",
            "",
        ]
    )
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT} sku={len(by_sku)} names={len(by_name)} items_with_q={len(items)}")
    for q, n in samples[:12]:
        print(f"  {q:4d} | {n[:90]}")


if __name__ == "__main__":
    main()
