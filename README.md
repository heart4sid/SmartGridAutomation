# SmartGridSoft Automation Portal

> **Automated resident data scraping** from [smartgridsoft.in](https://www.smartgridsoft.in) — extracts Profile & Meter data for thousands of accounts with a single click.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| **Batch Scraping** | Upload an Excel file with credentials, scrape all accounts automatically |
| **Resume Support** | Cancelled, paused, or failed jobs can be **resumed** from where they left off — no re-scraping |
| **Auto-Pause** | If the target website goes offline, the job pauses automatically and can be resumed later |
| **Partial Exports** | Download Excel/CSV/JSON exports even from incomplete jobs |
| **Real-Time Dashboard** | Live progress bar, metrics, and activity log updated every 2 seconds |
| **Job History** | Browse and reload any previous job from the built-in history panel |
| **Multi-Format Export** | Excel (multi-sheet), CSV, JSON, and Failed Logins CSV |
| **Smart Retry** | Each account is retried up to 3 times on transient failures |
| **Headless Browser** | Uses Playwright (Chromium) — no visible browser window needed |

---

## 🏗️ Architecture

```
SmartGridAutomation/
├── main.py          # FastAPI server — APIs + serves frontend
├── worker.py        # Playwright scraper — runs as subprocess
├── database.py      # SQLAlchemy models (SQLite)
├── requirements.txt # Python dependencies
└── static/
    └── index.html   # Self-contained frontend (HTML + CSS + JS)
```

**How it works:**

1. User uploads an Excel file with `Username` and `Password` columns via the web portal
2. The server creates a Job in the SQLite database and spawns a worker subprocess
3. The worker uses Playwright to login to each account, scrape Profile & MyMeter pages
4. Progress is stored in the database; the frontend polls for updates every 2 seconds
5. On completion (or cancellation/pause), export files are generated automatically

---

## 🚀 Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt
playwright install chromium

# 2. Start the server
python main.py

# 3. Open in browser
# → http://127.0.0.1:8000
```

> 📖 For a detailed step-by-step guide, see [GETTING_STARTED.md](GETTING_STARTED.md)

---

## 📡 API Reference

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/jobs` | `POST` | Create a new scraping job (multipart form: `society_name` + `file`) |
| `/api/jobs` | `GET` | List all jobs (newest first) |
| `/api/jobs/{id}/status` | `GET` | Get detailed job status with per-record logs |
| `/api/jobs/{id}/cancel` | `POST` | Cancel a running job |
| `/api/jobs/{id}/resume` | `POST` | Resume a cancelled/paused/failed job |
| `/api/jobs/{id}/export` | `POST` | Force-generate exports for partial data |
| `/api/jobs/{id}/download/{fmt}` | `GET` | Download export file (`excel`, `csv`, `json`, `failed`) |

---

## 🛠️ Technology Stack

- **Backend:** Python, FastAPI, SQLAlchemy, SQLite
- **Scraper:** Playwright (async Chromium)
- **Frontend:** Vanilla HTML/CSS/JS (single-file, no build step)
- **Font:** Inter (Google Fonts)

---

## 📋 Supported Societies

Bharat City, Olive County, Gulshan Ikebana, Exotica Dreamville, Orbit Plaza, Cherry County, Shalimar City, Nirala Estate, Bharat City Phase-2, JYOTI SUPER VILLAGE, AJNARA HOMES, RAJNAGAR REIDENCY, Cleo County, RAJ Empire, Addela, Windsor Court, Ajnara Daffodil, Sukh Sagar Complex, Hawelia Valencia Homes, Palm Resort, Shalimar City RG, Hawelia Valenova Park, Green Arch, COCO County, County 107, IVY County

---

## 📄 License

Private / Internal Use Only
