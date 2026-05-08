# DRAP Daily Update Monitor

A Python monitoring tool for tracking public updates from the Drug Regulatory Authority of Pakistan (DRAP) website.

The script checks configured public DRAP pages, discovers relevant internal links and documents, classifies updates by keyword, and generates daily machine-readable and human-readable reports.

## Features

- Crawls configured public pages at a low request rate
- Supports WordPress API collection when enabled
- Respects `robots.txt`
- Detects blocked/CAPTCHA pages and records a manual-check status
- Classifies updates as good, bad, or neutral using configurable keywords
- Exports daily JSON and CSV files
- Generates DOCX and PDF reports
- Uses a development cache to avoid repeated page fetches while testing

## Tech Stack

- Python
- Scrapling
- python-docx
- ReportLab
- Standard-library HTML parsing, CSV, JSON, and URL tools

## Project Structure

```text
drap-daily-update-monitor/
  drap_monitor.py
  config.example.json
  requirements.txt
  README.md
  .gitignore
```

Generated folders such as `data/`, `reports/`, `.cache/`, `.venv/`, and `__pycache__/` are intentionally excluded from GitHub.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
Copy-Item config.example.json config.json
```

Edit `config.json` if you want to change the starting URLs, keywords, request delay, crawl limit, or report behavior.

## Run

```powershell
python drap_monitor.py
```

Reports are written to:

```text
reports/
```

Machine-readable output is written to:

```text
data/
```

## Configuration Highlights

Useful settings in `config.json`:

```json
{
  "max_pages_per_run": 500,
  "report_scope": "new_only",
  "record_all_internal_links": true,
  "request_delay_seconds": 1,
  "robots_txt_obey": true
}
```

`report_scope` supports:

- `new_only`: report only items not previously seen
- `all_found_this_run`: report every discovered item from the current run

## Responsible Use

This project is configured for public, low-rate monitoring. It does not use proxy rotation, stealth bypass, or automated CAPTCHA solving. If a blocking or verification page is detected, the run records `manual check needed`.
