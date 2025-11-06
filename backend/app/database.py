from sqlmodel import SQLModel, create_engine, Session
from pathlib import Path
import sqlite3

# مسار قاعدة البيانات
DB_PATH = (Path(__file__).resolve().parent.parent / "workshop.db")
DATABASE_URL = f"sqlite:///{DB_PATH}"

# محرك SQLite (خيط واحد للويب + WAL)
engine = create_engine(
    DATABASE_URL,
    echo=False,
    connect_args={"check_same_thread": False}
)

def _enable_wal():
    if not DB_PATH.exists():
        return
    con = sqlite3.connect(str(DB_PATH))
    try:
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA foreign_keys=ON")
        con.commit()
    finally:
        con.close()

def _create_indexes():
    con = sqlite3.connect(str(DB_PATH))
    try:
        cur = con.cursor()
        # فهارس بحث سريع
        cur.execute('CREATE INDEX IF NOT EXISTS idx_enginesupply_serial ON enginesupply(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_engineissue_serial ON engineissue(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_enginerehab_serial ON enginerehab(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_enginecheck_serial ON enginecheck(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_engineupload_serial ON engineupload(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_enginelathe_serial ON enginelathe(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_enginepump_serial ON enginepump(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_engineelectrical_serial ON engineelectrical(serial)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_generatorsupply_code ON generatorsupply(code)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_generatorissue_code ON generatorissue(code)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_generatorinspect_code ON generatorinspect(code)')
        con.commit()
    finally:
        con.close()

def init_db():
    from . import models  # للتسجيل
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    SQLModel.metadata.create_all(engine)
    _enable_wal()
    _create_indexes()

def get_session():
    with Session(engine) as session:
        yield session
