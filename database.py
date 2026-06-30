"""
SmartGridAutomation — Database Models
SQLAlchemy models for job tracking and scraped resident data.
"""
import os
from datetime import datetime, timezone
from sqlalchemy import create_engine, Column, String, Integer, DateTime, Text, ForeignKey
from sqlalchemy.orm import declarative_base, sessionmaker, relationship

# Database file lives in the project root
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_URL = f"sqlite:///{os.path.join(BASE_DIR, 'automation.db')}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Job(Base):
    """Represents a single automation run (one Excel upload)."""
    __tablename__ = "jobs"

    id = Column(String, primary_key=True, index=True)
    status = Column(String, default="PENDING")  # PENDING, PROCESSING, COMPLETED, FAILED, CANCELLED
    society_name = Column(String)
    total_records = Column(Integer, default=0)
    processed_records = Column(Integer, default=0)
    successful = Column(Integer, default=0)
    failed = Column(Integer, default=0)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    completed_at = Column(DateTime, nullable=True)
    # Paths to generated export files
    export_excel = Column(String, nullable=True)
    export_csv = Column(String, nullable=True)
    export_json = Column(String, nullable=True)
    export_failed_csv = Column(String, nullable=True)

    residents = relationship("ResidentData", back_populates="job", cascade="all, delete-orphan")


class ResidentData(Base):
    """
    Stores ALL scraped data for one resident as a JSON blob.
    This avoids needing 50+ columns — the full structured data
    (profile, meter, rates, recharges, etc.) is stored in `data_json`.
    """
    __tablename__ = "resident_data"

    id = Column(Integer, primary_key=True, autoincrement=True)
    job_id = Column(String, ForeignKey("jobs.id"), index=True)
    username = Column(String)          # What was entered as username (e.g. "A1")
    password = Column(String)          # What was entered as password (e.g. "101")
    status = Column(String)            # Success, Auth Failed, Error, Pending
    error_message = Column(Text, nullable=True)
    data_json = Column(Text, nullable=True)  # Full scraped data as JSON string
    scraped_at = Column(DateTime, nullable=True)

    job = relationship("Job", back_populates="residents")


def init_db():
    """Create all tables if they don't exist."""
    Base.metadata.create_all(bind=engine)


def get_db():
    """FastAPI dependency for database sessions."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
