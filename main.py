"""
SmartGridAutomation — FastAPI Server
Single-server app: serves both the API and the frontend HTML.
Run with: python main.py
"""
import sys
import asyncio

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

import os
import io
import uuid
import json
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, HTTPException, Depends
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from sqlalchemy.orm import Session

from database import init_db, get_db, Job, ResidentData
from worker import run_automation_job, generate_exports

# ---------------------------------------------------------------------------
# App Setup
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent

app = FastAPI(title="SmartGridSoft Automation Portal")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def startup():
    init_db()


# ---------------------------------------------------------------------------
# Serve Frontend
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    """Serve the self-contained frontend HTML."""
    index_path = BASE_DIR / "static" / "index.html"
    if index_path.exists():
        return HTMLResponse(content=index_path.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>Frontend not found. Place index.html in static/ folder.</h1>")


# ---------------------------------------------------------------------------
# API: Create Job
# ---------------------------------------------------------------------------
@app.post("/api/jobs")
async def create_job(
    background_tasks: BackgroundTasks,
    society_name: str = Form(...),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    """
    Accept Excel file + society name, parse credentials, create job, dispatch worker.
    Smart column detection: treats Col 1 as Username, Col 2 as Password.
    """
    # Validate extension
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(400, "Only Excel files (.xlsx or .xls) are supported.")

    # Parse Excel
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), read_only=True)
        sheet = wb.active

        rows_raw = list(sheet.iter_rows(values_only=True))
        wb.close()

        if len(rows_raw) < 2:
            raise HTTPException(400, "Excel file must have at least a header row and one data row.")

        # Smart detection: check if first row looks like a header
        first_row = rows_raw[0]
        start_idx = 0

        if first_row and len(first_row) >= 2:
            val1 = str(first_row[0]).strip() if first_row[0] else ""
            val2 = str(first_row[1]).strip() if first_row[1] else ""
            # If both values are purely text (not numeric), treat as header and skip
            is_header = False
            for keyword in ["username", "password", "user", "pass", "tower", "flat",
                            "customer", "id", "name", "credential", "login"]:
                if keyword in val1.lower() or keyword in val2.lower():
                    is_header = True
                    break
            # Also skip if both values look like column identifiers (single letters/short text)
            if not is_header and not val2.replace(".", "").isdigit():
                # Second column is not numeric — likely a header
                is_header = True
            if is_header:
                start_idx = 1

        # Parse data rows
        credentials = []
        for row in rows_raw[start_idx:]:
            if not row or len(row) < 2:
                continue
            username = str(row[0]).strip() if row[0] is not None else ""
            password = str(row[1]).strip() if row[1] is not None else ""
            # Clean up float passwords like "101.0" → "101"
            if password.endswith(".0"):
                password = password[:-2]
            if username and username.lower() not in ["nan", "none", ""]:
                credentials.append((username, password))

        if not credentials:
            raise HTTPException(400, "No valid credential rows found in the Excel file.")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"Failed to parse Excel: {str(e)}")

    # Create Job
    job_id = str(uuid.uuid4())
    job = Job(
        id=job_id,
        status="PENDING",
        society_name=society_name,
        total_records=len(credentials),
        processed_records=0,
        successful=0,
        failed=0,
    )
    db.add(job)

    for username, password in credentials:
        record = ResidentData(
            job_id=job_id,
            username=username,
            password=password,
            status="Pending",
        )
        db.add(record)

    db.commit()

    # Dispatch background worker as a separate OS process to prevent event loop policy crashes
    import subprocess
    worker_script = str(BASE_DIR / "worker.py")
    subprocess.Popen([sys.executable, worker_script, job_id], close_fds=True)

    return {
        "job_id": job_id,
        "message": "Automation started",
        "total_records": len(credentials),
    }


# ---------------------------------------------------------------------------
# API: Job Status
# ---------------------------------------------------------------------------
@app.get("/api/jobs/{job_id}/status")
async def get_job_status(job_id: str, db: Session = Depends(get_db)):
    """Return full job progress with per-record details."""
    job = db.query(Job).filter(Job.id == job_id).first()
    if not job:
        raise HTTPException(404, "Job not found.")

    records = db.query(ResidentData).filter(ResidentData.job_id == job_id).all()

    logs = []
    for r in records:
        entry = {
            "username": r.username,
            "password": r.password,
            "status": r.status,
            "error": r.error_message,
            "timestamp": r.scraped_at.strftime("%H:%M:%S") if r.scraped_at else "",
        }
        # Include key scraped data in the log for display
        if r.data_json:
            try:
                data = json.loads(r.data_json)
                profile = data.get("profile", {})
                meter = data.get("meter", {})
                entry["name"] = profile.get("name", "")
                entry["balance"] = meter.get("balance", "")
                entry["eb_units"] = meter.get("eb_units", "")
                entry["meter_status"] = meter.get("meter_status", "")
            except Exception:
                pass
        logs.append(entry)

    return {
        "job_id": job.id,
        "status": job.status,
        "society_name": job.society_name,
        "total_records": job.total_records,
        "processed_records": job.processed_records,
        "successful": job.successful,
        "failed": job.failed,
        "has_exports": job.export_excel is not None,
        "logs": logs,
    }


# ---------------------------------------------------------------------------
# API: Download Exports
# ---------------------------------------------------------------------------
@app.get("/api/jobs/{job_id}/download/{fmt}")
async def download_export(job_id: str, fmt: str, db: Session = Depends(get_db)):
    """Download export files. fmt = excel | csv | json | failed"""
    job = db.query(Job).filter(Job.id == job_id).first()
    if not job:
        raise HTTPException(404, "Job not found.")

    path_map = {
        "excel": (job.export_excel, "Residents.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "csv": (job.export_csv, "Residents.csv", "text/csv"),
        "json": (job.export_json, "Residents.json", "application/json"),
        "failed": (job.export_failed_csv, "failed_logins.csv", "text/csv"),
    }

    if fmt not in path_map:
        raise HTTPException(400, f"Invalid format: {fmt}. Use: excel, csv, json, failed")

    file_path, filename, media_type = path_map[fmt]
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(400, "Export file not available yet.")

    return FileResponse(path=file_path, filename=filename, media_type=media_type)


# ---------------------------------------------------------------------------
# API: Cancel Job
# ---------------------------------------------------------------------------
@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str, db: Session = Depends(get_db)):
    """Cancel a running job."""
    job = db.query(Job).filter(Job.id == job_id).first()
    if not job:
        raise HTTPException(404, "Job not found.")

    if job.status in ["PENDING", "PROCESSING"]:
        job.status = "CANCELLED"
        db.commit()
        return {"message": "Job cancellation requested."}

    return {"message": f"Job is already {job.status}."}


# ---------------------------------------------------------------------------
# API: Resume Job
# ---------------------------------------------------------------------------
@app.post("/api/jobs/{job_id}/resume")
async def resume_job(job_id: str, db: Session = Depends(get_db)):
    """
    Resume a cancelled, paused, or failed job.
    Re-launches the worker subprocess which will skip already-scraped records.
    """
    job = db.query(Job).filter(Job.id == job_id).first()
    if not job:
        raise HTTPException(404, "Job not found.")

    if job.status not in ["CANCELLED", "PAUSED", "FAILED"]:
        raise HTTPException(400, f"Cannot resume a job with status '{job.status}'. Only CANCELLED, PAUSED, or FAILED jobs can be resumed.")

    # Count remaining records
    pending_count = db.query(ResidentData).filter(
        ResidentData.job_id == job_id,
        ResidentData.status.in_(["Pending", "Error"])
    ).count()

    if pending_count == 0:
        raise HTTPException(400, "No pending records left to process. All records have been scraped.")

    # Reset status — the worker will set it to PROCESSING
    job.status = "PENDING"
    db.commit()

    # Spawn new worker subprocess
    import subprocess
    worker_script = str(BASE_DIR / "worker.py")
    subprocess.Popen([sys.executable, worker_script, job_id], close_fds=True)

    return {
        "message": f"Job resumed. {pending_count} remaining accounts will be processed.",
        "job_id": job_id,
        "remaining_records": pending_count,
    }


# ---------------------------------------------------------------------------
# API: Force Export (partial data)
# ---------------------------------------------------------------------------
@app.post("/api/jobs/{job_id}/export")
async def force_export(job_id: str, db: Session = Depends(get_db)):
    """
    Force-generate exports for any job that has at least 1 successful record.
    Useful for cancelled/failed jobs that weren't auto-exported.
    """
    job = db.query(Job).filter(Job.id == job_id).first()
    if not job:
        raise HTTPException(404, "Job not found.")

    if job.status == "PROCESSING":
        raise HTTPException(400, "Cannot export while job is still running.")

    success_count = db.query(ResidentData).filter(
        ResidentData.job_id == job_id,
        ResidentData.status == "Success"
    ).count()

    if success_count == 0:
        raise HTTPException(400, "No successful records to export.")

    try:
        exports = generate_exports(job_id, job.society_name, db)
        job.export_excel = exports["excel"]
        job.export_csv = exports["csv"]
        job.export_json = exports["json"]
        job.export_failed_csv = exports["failed_csv"]
        db.commit()
        return {
            "message": f"Exports generated with {success_count} successful records.",
            "records_exported": success_count,
        }
    except Exception as e:
        raise HTTPException(500, f"Failed to generate exports: {str(e)}")


# ---------------------------------------------------------------------------
# API: Job History
# ---------------------------------------------------------------------------
@app.get("/api/jobs")
async def list_jobs(db: Session = Depends(get_db)):
    """Return all jobs ordered by creation date (newest first)."""
    jobs = db.query(Job).order_by(Job.created_at.desc()).all()
    return [
        {
            "job_id": j.id,
            "status": j.status,
            "society_name": j.society_name,
            "total_records": j.total_records,
            "processed_records": j.processed_records,
            "successful": j.successful,
            "failed": j.failed,
            "has_exports": j.export_excel is not None,
            "created_at": j.created_at.isoformat() if j.created_at else "",
            "completed_at": j.completed_at.isoformat() if j.completed_at else "",
        }
        for j in jobs
    ]


# ---------------------------------------------------------------------------
# Run Server
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    print("\n" + "=" * 60)
    print("  SmartGridSoft Automation Portal")
    print("  Open in browser: http://127.0.0.1:8000")
    print("=" * 60 + "\n")
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)

