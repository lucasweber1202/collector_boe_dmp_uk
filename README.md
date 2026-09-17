# collector_boe_dmp_uk

Standalone collector for 11 official aggregate Bank of England Decision Maker Panel series: realised/expected own-price, wage, employment and unit-cost growth plus CPI expectations. It stores 799 monthly observations from 2017-01 through 2026-08; restricted microdata are not used.

The collector probes recent official monthly release pages because the former DMP index URL no longer exists. Survey month is `reference_date`; public release date is `available_at` for newest observations. Older values without archived timestamps are `first_seen`.

## Install and run (PowerShell)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
Copy-Item .env.example .env
pytest -q
python main.py
```

Set `COLLECTOR_DB_URL` and allow `bankofengland.co.uk`. Databricks is optional via `.[databricks]`. Source smoke: `python -c "from scripts.extract import collect; x=collect(); print(len(x.catalog), len(x.observations))"`.

See [METHODOLOGY.md](METHODOLOGY.md) and [POINT_IN_TIME.md](POINT_IN_TIME.md).
