# AGENTS.md

## Cursor Cloud specific instructions

This repo is a Python purchase/procurement forecasting tool for 1С Корп Общепит.
Offline Excel-in / Excel-out: upload a stock (`остатки`) file and a sales
(`продажи`) file exported from 1C, get back a multi-sheet Excel workbook
(order forecast, ABC analysis, OOS/dead-stock/overstock risks). There is no
database, network, or external API — all processing is in-memory on the
uploaded spreadsheets.

### Environment
- Python deps live in a virtualenv at `.venv` (created by the startup update
  script). Activate with `source .venv/bin/activate`, or call binaries
  directly, e.g. `.venv/bin/python`, `.venv/bin/streamlit`.
- The desktop GUI needs system Tk (`python3-tk`) and `customtkinter`; these are
  provided in the snapshot / `requirements-desktop.txt`.

### Package layout (non-obvious)
- All modules live in packages (`calculations/`, `config/`, `data/`, `excel/`,
  `utils/`, `gui/`, `samples/`), and every module imports via those package
  paths. Run entry points from the repo root so imports resolve. If you add a
  new module, put it in the correct package (do not flatten files into root).

### Services / entry points (all run from repo root)
- Streamlit web app (primary): `streamlit run streamlit_app.py`
  (defaults to port 8501; use `--server.port <N>` to change). This is the
  best way to test end to end headlessly.
- Headless CLI: `python run_cli.py --stock <xlsx> --sales <xlsx> --from DD.MM.YYYY --to DD.MM.YYYY --order-days 14 [--order-coef 1.0] [--out out.xlsx]`
- Desktop GUI: `python main.py` — customtkinter/Tk window; requires a graphical
  display, so it is NOT runnable in a headless VM. Prefer Streamlit or the CLI.

### Demo / test data
- Generate demo Excel files: `python samples/create_samples.py`
  (writes `samples/остатки_демо.xlsx` and `samples/продажи_демо.xlsx`).
- Regenerate blank user templates: `python samples/create_templates.py`
  (writes into `samples/templates/`).
- Loaders read only the `Данные` sheet; the `Инструкция` sheet is ignored.
- Date parsing uses `dayfirst=True`, so ISO (`YYYY-MM-DD`) demo dates can spread
  across the year. Pick a wide calculation period (e.g. full year) when using
  the generated demo data, or the period filter may catch few rows.

### Output
- Results are written to `output/` (auto-created); files are named
  `Заказ_прогноз_<timestamp>.xlsx`.

### Tests / lint / build
- There are no automated tests, no lint config, and no build step in this repo.
  Validate changes by running the CLI or Streamlit app end to end and opening
  the produced workbook.
