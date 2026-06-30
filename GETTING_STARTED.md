# Getting Started — SmartGridSoft Automation Portal

This guide walks you through everything step by step. **No programming knowledge required.**

---

## 📋 Prerequisites

You need the following installed on your computer:

### 1. Python (version 3.9 or higher)

**Check if you have it:**
```
python --version
```

**If not installed:** Download from [python.org](https://www.python.org/downloads/) and install. During installation, make sure to check ✅ **"Add Python to PATH"**.

---

## 🔧 Installation

### Step 1: Open a Terminal

- Press `Win + R`, type `cmd`, and press Enter
- Or search for "Command Prompt" or "PowerShell" in the Start Menu

### Step 2: Navigate to the Project Folder

```
cd "E:\Ritik Project\SmartGridAutomation"
```

### Step 3: Install Python Dependencies

```
pip install -r requirements.txt
```

This installs: FastAPI, Uvicorn, Playwright, SQLAlchemy, and openpyxl.

### Step 4: Install the Browser Engine

```
playwright install chromium
```

This downloads a headless Chromium browser that the scraper uses behind the scenes. It only needs to be done once.

---

## 🚀 Running the Application

### Start the Server

```
python main.py
```

You should see:
```
============================================================
  SmartGridSoft Automation Portal
  Open in browser: http://127.0.0.1:8000
============================================================
```

### Open in Your Browser

Go to: **http://127.0.0.1:8000**

You'll see the automation portal with a dark, modern interface.

---

## 📊 Preparing Your Excel File

Your Excel file should have **two columns**:

| Column 1 | Column 2 |
|-----------|----------|
| Username  | Password |
| A1        | 101      |
| A1        | 102      |
| B3        | 201      |
| ...       | ...      |

**Rules:**
- The file must be `.xlsx` or `.xls` format
- The first row can be a header (it will be automatically detected and skipped)
- Column 1 = Username/Tower code, Column 2 = Password/Flat number
- Empty rows are automatically skipped

---

## 🖥️ Using the Portal

### Starting a New Job

1. **Select a Society** — Type in the search box and click to select
2. **Upload Excel File** — Drag and drop, or click the upload area
3. **Click "Start Automation"** — The button activates when both fields are filled

### During Execution

- Watch the **real-time progress bar** and **metrics** (Total, Processed, Success, Failed)
- The **Activity Log** table shows each account's result as it's processed
- You can **Stop Execution** at any time using the red button

### After Completion

- **Download buttons** appear for Excel, CSV, JSON, and Failed Logins
- Click **"New Automation"** to start a fresh run

---

## 🔄 Resuming a Cancelled/Failed Job

This is the most important feature for large batches (hundreds or thousands of accounts):

1. If a job is **cancelled**, **paused** (website went offline), or **failed**, you'll see a **"Resume Job"** button
2. Click **"Resume Job"** — the scraper picks up exactly where it left off
3. Already-scraped accounts are **never re-scraped** — only pending accounts are processed
4. Partial export files are **automatically generated** with whatever data was collected so far

### Using Job History

- Click the **"📋 Job History"** toggle at the bottom of the Control Center
- See all previous jobs with their status and progress
- Click any job to load its status and resume/export it

---

## ⚠️ Troubleshooting

### "Connection Timed Out" errors

The target website (smartgridsoft.in) may be temporarily offline. The scraper will:
1. Retry each account up to **3 times**
2. After 3 consecutive connection failures, **wait 60 seconds** and check again
3. If still offline, the job **pauses** automatically
4. You can **resume** it later when the website is back

### Job was cancelled mid-way

No data is lost! Click **"Resume Job"** to continue from where it stopped. All previously scraped data is safe in the database.

### Server won't start / Port already in use

If you see an error about port 8000 being in use:
1. Close any other terminals running `python main.py`
2. Or change the port: edit `main.py` and change `port=8000` to `port=8001`

### Excel file not accepted

Make sure:
- File extension is `.xlsx` or `.xls`
- File has at least 2 columns (Username, Password)
- File has at least 1 data row (not counting the header)

---

## 📁 Output Files

After a job completes (or is partially completed), exports are saved in:

```
exports/{job-id-prefix}/
├── Residents.xlsx     # Multi-sheet Excel workbook
├── Residents.csv      # Flat CSV with key fields
├── Residents.json     # Full structured JSON data
└── failed_logins.csv  # Accounts that failed authentication
```

### Excel Workbook Sheets

| Sheet | Contents |
|-------|----------|
| **Summary** | All profile + meter data in a flat table |
| **Monthly Consumption** | Month-by-month usage history |
| **Recharge History** | Last 10 recharges per account |
| **Current Month Consumption** | Daily readings for the current month |

---

## 🔒 Important Notes

- **Credentials are stored locally** in a SQLite database (`automation.db`) on your computer only
- **No data is sent anywhere** except to the target website for login
- The scraper runs a **headless browser** — no visible window opens
- **Back up your exports** — the database may be cleared between sessions

---

## 💡 Tips for Large Batches

1. **Stable internet** — Use a wired connection if possible
2. **Don't close the terminal** — The server must stay running while a job is in progress
3. **Use Resume** — If something goes wrong, always resume instead of restarting from scratch
4. **Download partial exports** — Even if a job is incomplete, download what you have so far
5. **Run during off-peak hours** — The target website may be slower during peak times
