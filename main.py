"""
Точка входа: модуль прогноза закупок для 1С Корп Общепит.

Запуск из корня проекта:
    python main.py
"""

from __future__ import annotations

import sys
from pathlib import Path

# Гарантируем, что корень проекта в sys.path (удобно для PyCharm)
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.logging_config import setup_logging


def main() -> int:
    setup_logging()
    from gui.app import run_app

    run_app()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
