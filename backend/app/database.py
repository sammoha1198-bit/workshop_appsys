# app/database.py
from __future__ import annotations
import os
from pathlib import Path
from typing import Generator

from sqlmodel import SQLModel, create_engine, Session
from sqlalchemy import text

# لو لديك DATABASE_URL من Render/Neon سيتم استعماله، وإلا نرجع لـ SQLite محليًا
SQLITE_PATH = (Path(__file__).resolve().parent.parent / "workshop.db")
DB_PATH = SQLITE_PATH  # legacy alias for old imports

DATABASE_URL = os.getenv("DATABASE_URL") or f"sqlite:///{SQLITE_PATH}"

is_sqlite = DATABASE_URL.startswith("sqlite")

# إنشاء المحرك
if is_sqlite:
    engine = create_engine(
        DATABASE_URL,
        echo=False,
        connect_args={"check_same_thread": False},  # خاص بـ SQLite
    )
else:
    # PostgreSQL (Neon)
    engine = create_engine(
        DATABASE_URL,
        echo=False,
        pool_pre_ping=True,
        pool_recycle=1800,
    )

def _create_indexes() -> None:
    """
    إنشاء الفهارس لكل الجداول لتسريع البحث سواء على PostgreSQL أو SQLite.
    نستخدم SQL قياسي مع IF NOT EXISTS (مدعوم في Postgres ≥ 9.5 و SQLite حديثة).
    """
    index_sql = [
        # Engines
        "CREATE INDEX IF NOT EXISTS idx_enginesupply_serial     ON enginesupply(serial)",
        "CREATE INDEX IF NOT EXISTS idx_engineissue_serial      ON engineissue(serial)",
        "CREATE INDEX IF NOT EXISTS idx_enginerehab_serial      ON enginerehab(serial)",
        "CREATE INDEX IF NOT EXISTS idx_enginecheck_serial      ON enginecheck(serial)",
        "CREATE INDEX IF NOT EXISTS idx_engineupload_serial     ON engineupload(serial)",
        "CREATE INDEX IF NOT EXISTS idx_enginelathe_serial      ON enginelathe(serial)",
        "CREATE INDEX IF NOT EXISTS idx_enginepump_serial       ON enginepump(serial)",
        "CREATE INDEX IF NOT EXISTS idx_engineelectrical_serial ON engineelectrical(serial)",
        # Generators
        "CREATE INDEX IF NOT EXISTS idx_generatorsupply_code    ON generatorsupply(code)",
        "CREATE INDEX IF NOT EXISTS idx_generatorissue_code     ON generatorissue(code)",
        "CREATE INDEX IF NOT EXISTS idx_generatorinspect_code   ON generatorinspect(code)",
    ]
    with engine.begin() as conn:
        for sql in index_sql:
            conn.execute(text(sql))

def init_db() -> None:
    """إنشاء الجداول والفهارس."""
    from . import models  # تأكد من تسجيل النماذج
    if is_sqlite:
        SQLITE_PATH.parent.mkdir(parents=True, exist_ok=True)
    SQLModel.metadata.create_all(engine)
    _create_indexes()

def get_session() -> Generator[Session, None, None]:
    with Session(engine) as session:
        yield session
