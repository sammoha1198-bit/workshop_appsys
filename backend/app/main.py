# app/main.py
from fastapi import FastAPI, Body, Depends, Request, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from typing import Dict, Any, List, Union, Optional
from io import BytesIO
from urllib.parse import quote
import datetime as dt
import os
import sqlite3

from sqlmodel import Session, select
from app.database import init_db, get_session, engine, DB_PATH
from app import models

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = FastAPI(title="Alsami Workshop Backend", version="4.0-final")

# ---------------- CORS ----------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://workshop-frontend-cmqd.onrender.com"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------- Utilities ----------------
def _now_iso() -> str:
    return dt.datetime.utcnow().isoformat()

def _pragma_table_info(conn: sqlite3.Connection, table: str) -> List[str]:
    cur = conn.execute(f'PRAGMA table_info("{table}")')
    return [row[1] for row in cur.fetchall()]  # column names

def _ensure_table(conn: sqlite3.Connection, table: str):
    """يضمن وجود الجدول (لو ما وُجد؛ يُنشأ سريعًا بالحد الأدنى)."""
    try:
        conn.execute(f'SELECT 1 FROM "{table}" LIMIT 1')
    except sqlite3.OperationalError:
        # ننشئ نسخة بسيطة من الجدول (مفتاح + حقول أساسية)؛
        # الجداول سيُكمّلها SQLModel.create_all في init_db لاحقًا.
        minimal = {
            "enginesupply":       'CREATE TABLE IF NOT EXISTS enginesupply (id INTEGER PRIMARY KEY, serial TEXT)',
            "engineissue":        'CREATE TABLE IF NOT EXISTS engineissue (id INTEGER PRIMARY KEY, serial TEXT)',
            "enginerehab":        'CREATE TABLE IF NOT EXISTS enginerehab (id INTEGER PRIMARY KEY, serial TEXT)',
            "enginecheck":        'CREATE TABLE IF NOT EXISTS enginecheck (id INTEGER PRIMARY KEY, serial TEXT)',
            "engineupload":       'CREATE TABLE IF NOT EXISTS engineupload (id INTEGER PRIMARY KEY, serial TEXT)',
            "enginelathe":        'CREATE TABLE IF NOT EXISTS enginelathe (id INTEGER PRIMARY KEY, serial TEXT)',
            "enginepump":         'CREATE TABLE IF NOT EXISTS enginepump (id INTEGER PRIMARY KEY, serial TEXT)',
            "engineelectrical":   'CREATE TABLE IF NOT EXISTS engineelectrical (id INTEGER PRIMARY KEY, serial TEXT)',

            "generatorsupply":    'CREATE TABLE IF NOT EXISTS generatorsupply (id INTEGER PRIMARY KEY, code TEXT)',
            "generatorissue":     'CREATE TABLE IF NOT EXISTS generatorissue (id INTEGER PRIMARY KEY, code TEXT)',
            "generatorinspect":   'CREATE TABLE IF NOT EXISTS generatorinspect (id INTEGER PRIMARY KEY, code TEXT)',
        }
        if table in minimal:
            conn.execute(minimal[table])
            conn.commit()

def _ensure_column(conn: sqlite3.Connection, table: str, column: str, coltype: str):
    _ensure_table(conn, table)
    cols = _pragma_table_info(conn, table)
    if column not in cols:
        conn.execute(f'ALTER TABLE "{table}" ADD COLUMN "{column}" {coltype}')
        conn.commit()

def _self_heal_schema():
    """
    يصلّح الأعمدة الناقصة بدون تدمير البيانات:
    - enginesupply.engineType
    - enginecheck.desc (بعض النسخ كانت 'description' فقط)
    ...إلخ
    """
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA foreign_keys=ON")

    try:
        # محركات - توريد
        _ensure_column(conn, "enginesupply", "serial", "TEXT")
        _ensure_column(conn, "enginesupply", "engineType", "TEXT")
        _ensure_column(conn, "enginesupply", "model", "TEXT")
        _ensure_column(conn, "enginesupply", "prevSite", "TEXT")
        _ensure_column(conn, "enginesupply", "supDate", "TEXT")
        _ensure_column(conn, "enginesupply", "supplier", "TEXT")
        _ensure_column(conn, "enginesupply", "notes", "TEXT")

        # محركات - صرف
        _ensure_column(conn, "engineissue", "serial", "TEXT")
        _ensure_column(conn, "engineissue", "currSite", "TEXT")
        _ensure_column(conn, "engineissue", "receiver", "TEXT")
        _ensure_column(conn, "engineissue", "requester", "TEXT")
        _ensure_column(conn, "engineissue", "issueDate", "TEXT")
        _ensure_column(conn, "engineissue", "notes", "TEXT")

        # محركات - تأهيل
        _ensure_column(conn, "enginerehab", "serial", "TEXT")
        _ensure_column(conn, "enginerehab", "rehabber", "TEXT")
        _ensure_column(conn, "enginerehab", "rehabType", "TEXT")
        _ensure_column(conn, "enginerehab", "rehabDate", "TEXT")
        _ensure_column(conn, "enginerehab", "notes", "TEXT")

        # محركات - فحص (desc/description)
        _ensure_column(conn, "enginecheck", "serial", "TEXT")
        _ensure_column(conn, "enginecheck", "inspector", "TEXT")
        _ensure_column(conn, "enginecheck", "checkDate", "TEXT")
        _ensure_column(conn, "enginecheck", "notes", "TEXT")

        cols_check = _pragma_table_info(conn, "enginecheck")
        if "desc" not in cols_check and "description" in cols_check:
            _ensure_column(conn, "enginecheck", "desc", "TEXT")
            conn.execute('UPDATE enginecheck SET "desc" = COALESCE("desc", description)')
            conn.commit()
        else:
            _ensure_column(conn, "enginecheck", "desc", "TEXT")

        # محركات - رفع/مخرطة/بمبات/كهرباء
        _ensure_column(conn, "engineupload", "serial", "TEXT")
        _ensure_column(conn, "engineupload", "rehabUp", "TEXT")
        _ensure_column(conn, "engineupload", "checkUp", "TEXT")
        _ensure_column(conn, "engineupload", "rehabUpDate", "TEXT")
        _ensure_column(conn, "engineupload", "checkUpDate", "TEXT")
        _ensure_column(conn, "engineupload", "notes", "TEXT")

        _ensure_column(conn, "enginelathe", "serial", "TEXT")
        _ensure_column(conn, "enginelathe", "lathe", "TEXT")
        _ensure_column(conn, "enginelathe", "latheDate", "TEXT")
        _ensure_column(conn, "enginelathe", "notes", "TEXT")

        _ensure_column(conn, "enginepump", "serial", "TEXT")
        _ensure_column(conn, "enginepump", "pumpSerial", "TEXT")
        _ensure_column(conn, "enginepump", "pumpRehab", "TEXT")
        _ensure_column(conn, "enginepump", "notes", "TEXT")

        _ensure_column(conn, "engineelectrical", "serial", "TEXT")
        _ensure_column(conn, "engineelectrical", "etype", "TEXT")
        _ensure_column(conn, "engineelectrical", "starter", "TEXT")
        _ensure_column(conn, "engineelectrical", "alternator", "TEXT")
        _ensure_column(conn, "engineelectrical", "edate", "TEXT")
        _ensure_column(conn, "engineelectrical", "notes", "TEXT")

        # مولدات
        _ensure_column(conn, "generatorsupply", "code", "TEXT")
        _ensure_column(conn, "generatorsupply", "gType", "TEXT")
        _ensure_column(conn, "generatorsupply", "model", "TEXT")
        _ensure_column(conn, "generatorsupply", "prevSite", "TEXT")
        _ensure_column(conn, "generatorsupply", "supDate", "TEXT")
        _ensure_column(conn, "generatorsupply", "supplier", "TEXT")
        _ensure_column(conn, "generatorsupply", "vendor", "TEXT")
        _ensure_column(conn, "generatorsupply", "notes", "TEXT")

        _ensure_column(conn, "generatorissue", "code", "TEXT")
        _ensure_column(conn, "generatorissue", "issueDate", "TEXT")
        _ensure_column(conn, "generatorissue", "receiver", "TEXT")
        _ensure_column(conn, "generatorissue", "requester", "TEXT")
        _ensure_column(conn, "generatorissue", "currSite", "TEXT")
        _ensure_column(conn, "generatorissue", "notes", "TEXT")

        _ensure_column(conn, "generatorinspect", "code", "TEXT")
        _ensure_column(conn, "generatorinspect", "inspector", "TEXT")
        _ensure_column(conn, "generatorinspect", "elecRehab", "TEXT")
        _ensure_column(conn, "generatorinspect", "rehabDate", "TEXT")
        _ensure_column(conn, "generatorinspect", "rehabUp", "TEXT")
        _ensure_column(conn, "generatorinspect", "checkUp", "TEXT")
        _ensure_column(conn, "generatorinspect", "notes", "TEXT")

    except Exception as e:
        print("⚠️ schema self-heal error:", e)
    finally:
        conn.close()

def _safe_list(objs):
    return [o.dict(exclude_none=True) for o in objs]

def ensure_schema(session: Session):
    """
    يضمن إنشاء الجداول (SQLModel) ثم يُصلح الأعمدة الناقصة (SQLite).
    """
    try:
        init_db()            # SQLModel.metadata.create_all(engine)
        _self_heal_schema()  # إصلاح أعمدة ناقصة/قديمة
    except Exception as e:
        print("⚠️ ensure_schema error:", e)

# ---------------- Startup ----------------
@app.on_event("startup")
def _on_startup():
    init_db()
    _self_heal_schema()

# ---------------- Health ----------------
@app.get("/api/health")
def health() -> Dict[str, Any]:
    return {"ok": True, "time": _now_iso()}

# ---------------- Seed ----------------
@app.get("/api/seed")
def seed(session: Session = Depends(get_session)):
    """
    يُعبّي قاعدة البيانات ببيانات تجريبية إذا كانت فاضية.
    """
    try:
        ensure_schema(session)
        has_rows = session.exec(select(models.EngineSupply)).first()
        if has_rows:
            return {"ok": True, "msg": "data-exists"}
        # نستدعي seed من الملف الخاص إن وُجد، وإلا نضيف أقل بيانات تجريبية
        try:
            from app.seed_demo_data import seed as do_seed
            do_seed(session)
        except Exception:
            # fallback سريع
            session.add(models.EngineSupply(serial="111", engineType="Deutz", model="F4L912",
                                            prevSite="المخزن", supDate="2025-10-30",
                                            supplier="Yemen Mobile", notes="توريد جديد"))
            session.add(models.GeneratorSupply(code="GEN001", gType="30kVA", model="FG Wilson",
                                               prevSite="المستودع", supDate="2025-10-30",
                                               supplier="Yemen Mobile", vendor="PowerMax", notes="جديد"))
            session.commit()
        return {"ok": True, "msg": "seeded"}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

# ---------------- Admin Repair/Rebuild ----------------
@app.post("/api/admin/repair")
def admin_repair():
    try:
        _self_heal_schema()
        return {"ok": True}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

@app.post("/api/admin/rebuild")
def admin_rebuild():
    try:
        init_db()
        _self_heal_schema()
        return {"ok": True, "msg": "rebuilt"}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

# ---------------- Search ----------------
@app.get("/api/search/{key}")
def search_item(key: str, session: Session = Depends(get_session)) -> Dict[str, Any]:
    """
    يعيد نتائج كاملة للرقم التسلسلي للمحركات (serial)
    أو الكود للمولدات (code).
    """
    ensure_schema(session)
    result: Dict[str, Any] = {
        "engines": { "supply": [], "issue": [], "rehab": [], "check": [], "upload": [],
                     "lathe": [], "pump": [], "electrical": [] },
        "generators": { "supply": [], "issue": [], "inspect": [] },
    }
    try:
        # محركات (serial = key)
        result["engines"]["supply"] = _safe_list(
            session.exec(select(models.EngineSupply).where(models.EngineSupply.serial == key)).all()
        )
        result["engines"]["issue"] = _safe_list(
            session.exec(select(models.EngineIssue).where(models.EngineIssue.serial == key)).all()
        )
        result["engines"]["rehab"] = _safe_list(
            session.exec(select(models.EngineRehab).where(models.EngineRehab.serial == key)).all()
        )
        result["engines"]["check"] = _safe_list(
            session.exec(select(models.EngineCheck).where(models.EngineCheck.serial == key)).all()
        )
        result["engines"]["upload"] = _safe_list(
            session.exec(select(models.EngineUpload).where(models.EngineUpload.serial == key)).all()
        )
        result["engines"]["lathe"] = _safe_list(
            session.exec(select(models.EngineLathe).where(models.EngineLathe.serial == key)).all()
        )
        result["engines"]["pump"] = _safe_list(
            session.exec(select(models.EnginePump).where(models.EnginePump.serial == key)).all()
        )
        result["engines"]["electrical"] = _safe_list(
            session.exec(select(models.EngineElectrical).where(models.EngineElectrical.serial == key)).all()
        )

        # مولدات (code = key)
        result["generators"]["supply"] = _safe_list(
            session.exec(select(models.GeneratorSupply).where(models.GeneratorSupply.code == key)).all()
        )
        result["generators"]["issue"] = _safe_list(
            session.exec(select(models.GeneratorIssue).where(models.GeneratorIssue.code == key)).all()
        )
        result["generators"]["inspect"] = _safe_list(
            session.exec(select(models.GeneratorInspect).where(models.GeneratorInspect.code == key)).all()
        )
        return result
    except Exception as e:
        return JSONResponse({"ok": False, "error": f"search-error: {e}"}, status_code=500)

# ---------------- Last 3 ----------------
@app.get("/api/last3/engines")
def last3_engines(session: Session = Depends(get_session)) -> Dict[str, Any]:
    ensure_schema(session)
    try:
        rows = session.exec(
            select(models.EngineSupply).order_by(models.EngineSupply.id.desc()).limit(3)
        ).all()
        return {"items": [{"serial": r.serial, "prevSite": r.prevSite or ""} for r in rows]}
    except Exception as e:
        return {"items": [], "warn": str(e)}

@app.get("/api/last3/generators")
def last3_generators(session: Session = Depends(get_session)) -> Dict[str, Any]:
    ensure_schema(session)
    try:
        rows = session.exec(
            select(models.GeneratorSupply).order_by(models.GeneratorSupply.id.desc()).limit(3)
        ).all()
        return {"items": [{"code": r.code, "prevSite": r.prevSite or ""} for r in rows]}
    except Exception as e:
        return {"items": [], "warn": str(e)}

# ---------------- Sync (IndexedDB) + حفظ مفرد وإرجاع نتيجة مدمجة ----------------
@app.post("/api/sync/batch")
def sync_batch(payload: Dict[str, Any] = Body(...), session: Session = Depends(get_session)):
    """
    payload = { items: [ {store: 'eng_supply' | 'gen_issue' | ... , payload: {..fields..}} ] }
    يتجاهل أي عنصر خاطئ ويكمل.
    بعد الحفظ يرجّع {"ok":True,"saved":..,"skipped":..}
    """
    ensure_schema(session)
    items: List[Dict[str, Any]] = payload.get("items", [])
    saved, skipped = 0, 0

    STORE_TO_MODEL = {
        # محركات
        "eng_supply": models.EngineSupply,
        "eng_issue": models.EngineIssue,
        "eng_rehab": models.EngineRehab,
        "eng_check": models.EngineCheck,
        "eng_upload": models.EngineUpload,
        "eng_lathe": models.EngineLathe,
        "eng_pump": models.EnginePump,
        "eng_electrical": models.EngineElectrical,
        # مولدات
        "gen_supply": models.GeneratorSupply,
        "gen_issue": models.GeneratorIssue,
        "gen_inspect": models.GeneratorInspect,
    }

    for it in items:
        try:
            store = it.get("store")
            data = it.get("payload", {}) or {}
            Model = STORE_TO_MODEL.get(store)
            if not Model:
                skipped += 1
                continue
            obj = Model(**data)
            session.add(obj)
            saved += 1
        except Exception:
            skipped += 1

    session.commit()
    return {"ok": True, "saved": saved, "skipped": skipped}

@app.post("/api/save-and-search")
def save_and_search(payload: Dict[str, Any] = Body(...), session: Session = Depends(get_session)):
    """
    لحفظ عنصر واحد ثم إعادة نتيجة بحث مدمجة مباشرة.
    payload = { store: 'eng_supply' | ... , payload: {...}, key: '111' | 'GEN001' }
    """
    ensure_schema(session)
    try:
        STORE_TO_MODEL = {
            "eng_supply": models.EngineSupply,
            "eng_issue": models.EngineIssue,
            "eng_rehab": models.EngineRehab,
            "eng_check": models.EngineCheck,
            "eng_upload": models.EngineUpload,
            "eng_lathe": models.EngineLathe,
            "eng_pump": models.EnginePump,
            "eng_electrical": models.EngineElectrical,
            "gen_supply": models.GeneratorSupply,
            "gen_issue": models.GeneratorIssue,
            "gen_inspect": models.GeneratorInspect,
        }
        store = payload.get("store")
        data = payload.get("payload", {}) or {}
        key = payload.get("key") or data.get("serial") or data.get("code")
        Model = STORE_TO_MODEL.get(store)
        if not Model:
            raise HTTPException(status_code=400, detail="invalid store")
        session.add(Model(**data))
        session.commit()
        # رجّع نتيجة البحث فورًا
        return search_item(str(key), session)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"save-and-search error: {e}")

# ---------------- Export XLSX (قديم + جديد) ----------------
@app.post("/api/export/xlsx")
def export_xlsx(payload: Dict[str, Any] = Body(...), session: Session = Depends(get_session)):
    """
    body:
    {
      "scope": "engines"|"generators"|"both",
      "date_from": "YYYY-MM-DD",   # اختياري
      "date_to":   "YYYY-MM-DD",   # اختياري
      "filename": "تقرير.xlsx"     # اختياري
    }

    التعديل:
    - تصدير كل الحركات لكل رقم في صف واحد.
    - للمحركات: توريد، صرف، تأهيل، فحص، رفع، مخرطة، بمبات/نوزلات، كهرباء.
    - للمولدات: توريد، صرف، فحص/رفع.
    """
    ensure_schema(session)

    scope = (payload.get("scope") or "both").lower()
    date_from = _safe_date(payload.get("date_from"))
    date_to = _safe_date(payload.get("date_to"))
    original_filename: str = payload.get("filename") or "report.xlsx"
    safe_filename = "report.xlsx"
    encoded_name = quote(original_filename)

    def in_range(date_str: Optional[str]) -> bool:
        """تحقق أن التاريخ داخل المدى (إن وجد)."""
        if not (date_from or date_to):
            return True
        d = _safe_date(date_str)
        if not d:
            return False
        if date_from and d < date_from:
            return False
        if date_to and d > date_to:
            return False
        return True

    # إعداد ملف الإكسل
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="2563eb")
    alt_fill = PatternFill("solid", fgColor="f1f5f9")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    header_font = Font(bold=True, color="FFFFFF", name="Arial")
    normal_font = Font(name="Arial")
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin", color="DDDDDD"),
        top=Side(style="thin", color="DDDDDD"),
        bottom=Side(style="thin", color="DDDDDD"),
    )

    # ===================== المحركات =====================
    ws_eng = None
    if scope in ("engines", "both"):
        ws_eng = wb.active
        ws_eng.title = "المحركات"
        ws_eng.sheet_view.rightToLeft = True

        eng_headers = [
            "الرقم التسلسلي",
            "بيانات التوريد",
            "بيانات الصرف",
            "بيانات التأهيل",
            "بيانات الفحص",
            "بيانات الرفع",
            "بيانات المخرطة",
            "بيانات البمبات والنوزلات",
            "بيانات الكهرباء",
        ]
        ws_eng.append(eng_headers)
        for i in range(1, len(eng_headers) + 1):
            c = ws_eng.cell(row=1, column=i)
            c.fill = header_fill
            c.font = header_font
            c.alignment = align_right
            c.border = thin_border
            ws_eng.column_dimensions[c.column_letter].width = 35

        supplies = session.exec(select(models.EngineSupply)).all()
        issues = session.exec(select(models.EngineIssue)).all()
        rehabs = session.exec(select(models.EngineRehab)).all()
        checks = session.exec(select(models.EngineCheck)).all()
        uploads = session.exec(select(models.EngineUpload)).all()
        lathes = session.exec(select(models.EngineLathe)).all()
        pumps = session.exec(select(models.EnginePump)).all()
        elecs = session.exec(select(models.EngineElectrical)).all()

        by_serial: Dict[str, Dict[str, str]] = {}

        def add(seg: str, serial: str, text: str, date_field: Optional[str] = None):
            if not serial or not text:
                return
            if date_field and not in_range(date_field):
                return
            row = by_serial.setdefault(serial, {"serial": serial})
            old = row.get(seg) or ""
            row[seg] = (old + " | " if old else "") + text

        for r in supplies:
            if not in_range(r.supDate):
                continue
            txt = f"نوع:{r.engineType or ''} موديل:{r.model or ''} موقع سابق:{r.prevSite or ''} تاريخ:{r.supDate or ''} مورد:{r.supplier or ''} {r.notes or ''}"
            add("supply", r.serial, txt, r.supDate)

        for r in issues:
            txt = f"موقع حالي:{r.currSite or ''} مستلم:{r.receiver or ''} طالب:{r.requester or ''} تاريخ:{r.issueDate or ''} {r.notes or ''}"
            add("issue", r.serial, txt, r.issueDate)

        for r in rehabs:
            txt = f"جهة:{r.rehabber or ''} نوع:{r.rehabType or ''} تاريخ:{r.rehabDate or ''} {r.notes or ''}"
            add("rehab", r.serial, txt, r.rehabDate)

        for r in checks:
            txt = f"فاحص:{r.inspector or ''} وصف:{r.description or ''} تاريخ:{r.checkDate or ''} {r.notes or ''}"
            add("check", r.serial, txt, r.checkDate)

        for r in uploads:
            txt = f"رفع تأهيل:{r.rehabUp or ''} رفع فحص:{r.checkUp or ''} تاريخ تأهيل:{r.rehabUpDate or ''} تاريخ فحص:{r.checkUpDate or ''} {r.notes or ''}"
            add("upload", r.serial, txt)

        for r in lathes:
            txt = f"مخرطة:{r.lathe or ''} تاريخ:{r.latheDate or ''} {r.notes or ''}"
            add("lathe", r.serial, txt, r.latheDate)

        for r in pumps:
            txt = f"بمب:{r.pumpSerial or ''} تأهيل:{r.pumpRehab or ''} {r.notes or ''}"
            add("pump", r.serial, txt)

        for r in elecs:
            txt = f"نوع:{r.etype or ''} سلف:{r.starter or ''} دينمو:{r.alternator or ''} تاريخ:{r.edate or ''} {r.notes or ''}"
            add("elec", r.serial, txt, r.edate)

        row_idx = 2
        for serial, segs in sorted(by_serial.items()):
            row = [
                serial,
                segs.get("supply", ""),
                segs.get("issue", ""),
                segs.get("rehab", ""),
                segs.get("check", ""),
                segs.get("upload", ""),
                segs.get("lathe", ""),
                segs.get("pump", ""),
                segs.get("elec", ""),
            ]
            ws_eng.append(row)
            for col in range(1, len(row) + 1):
                c = ws_eng.cell(row=row_idx, column=col)
                c.font = normal_font
                c.alignment = align_right
                c.border = thin_border
                c.fill = alt_fill if row_idx % 2 else white_fill
            row_idx += 1

    # ===================== المولدات =====================
    ws_gen = None
    if scope in ("generators", "both"):
        ws_gen = wb.create_sheet(title="المولدات") if ws_eng else wb.active
        ws_gen.title = "المولدات"
        ws_gen.sheet_view.rightToLeft = True

        gen_headers = [
            "كود المولد",
            "بيانات التوريد",
            "بيانات الصرف",
            "بيانات الفحص / الرفع",
        ]
        ws_gen.append(gen_headers)
        for i in range(1, len(gen_headers) + 1):
            c = ws_gen.cell(row=1, column=i)
            c.fill = header_fill
            c.font = header_font
            c.alignment = align_right
            c.border = thin_border
            ws_gen.column_dimensions[c.column_letter].width = 35

        sup = session.exec(select(models.GeneratorSupply)).all()
        iss = session.exec(select(models.GeneratorIssue)).all()
        ins = session.exec(select(models.GeneratorInspect)).all()

        by_code: Dict[str, Dict[str, str]] = {}

        def addg(seg: str, code: str, text: str, date_field: Optional[str] = None):
            if not code or not text:
                return
            if date_field and not in_range(date_field):
                return
            row = by_code.setdefault(code, {"code": code})
            old = row.get(seg) or ""
            row[seg] = (old + " | " if old else "") + text

        for r in sup:
            if not in_range(r.supDate):
                continue
            txt = f"سعة:{r.gType or ''} موديل:{r.model or ''} موقع سابق:{r.prevSite or ''} تاريخ:{r.supDate or ''} جهة:{r.supplier or ''} مورد:{r.vendor or ''} {r.notes or ''}"
            addg("supply", r.code, txt, r.supDate)

        for r in iss:
            txt = f"تاريخ:{r.issueDate or ''} مستلم:{r.receiver or ''} طالب:{r.requester or ''} موقع حالي:{r.currSite or ''} {r.notes or ''}"
            addg("issue", r.code, txt, r.issueDate)

        for r in ins:
            txt = f"مفتش:{r.inspector or ''} تأهيل كهربائي:{r.elecRehab or ''} تاريخ التأهيل:{r.rehabDate or ''} رفع تأهيل:{r.rehabUp or ''} رفع فحص:{r.checkUp or ''} {r.notes or ''}"
            addg("inspect", r.code, txt, r.rehabDate)

        row_idx = 2
        for code, segs in sorted(by_code.items()):
            row = [
                code,
                segs.get("supply", ""),
                segs.get("issue", ""),
                segs.get("inspect", ""),
            ]
            ws_gen.append(row)
            for col in range(1, len(row) + 1):
                c = ws_gen.cell(row=row_idx, column=col)
                c.font = normal_font
                c.alignment = align_right
                c.border = thin_border
                c.fill = alt_fill if row_idx % 2 else white_fill
            row_idx += 1

    if not ws_eng and not ws_gen:
        ws = wb.active
        ws.title = "فارغ"
        ws.append(["لا توجد بيانات للتصدير"])

    ws = ws_gen or ws_eng or wb.active
    last_row = ws.max_row + 2
    sig = ws.cell(row=last_row, column=1)
    sig.value = f"تاريخ التوليد: {datetime.datetime.utcnow().isoformat()}"
    sig.alignment = align_right
    sig.font = Font(italic=True, color="666666", name="Arial")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    headers_resp = {
        "Content-Disposition": f"attachment; filename={safe_filename}; filename*=UTF-8''{encoded_name}"
    }
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers_resp,
    )

# ---------------- Root ----------------
@app.get("/")
def root():
    return {"service": "Workshop API", "status": "running", "static": "/static/index.html"}

# ---------------- Favicon (لمنع 404 في Render) ----------------
@app.get("/favicon.ico")
def favicon():
    return JSONResponse(content={}, status_code=204)

# ---------------- Local run ----------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=9000, reload=True)
