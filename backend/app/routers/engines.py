from fastapi import APIRouter, Depends
from sqlmodel import Session, select
from datetime import datetime
from ..database import get_session
from ..models import (
    EngineSupply,
    EngineIssue,
    EngineRehab,
    EngineCheck,
    EngineUpload,
    EngineLathe,
    EnginePump,
    EngineElectrical,
)

router = APIRouter(prefix="/api/engines", tags=["engines"])

# ========== 1️⃣ التوريد ==========
@router.post("/supply")
def add_engine_supply(data: EngineSupply, session: Session = Depends(get_session)):
    data.date = data.date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/supply")
def list_engine_supply(session: Session = Depends(get_session)):
    return session.exec(select(EngineSupply)).all()


# ========== 2️⃣ الصرف ==========
@router.post("/issue")
def add_engine_issue(data: EngineIssue, session: Session = Depends(get_session)):
    data.date = data.date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/issue")
def list_engine_issue(session: Session = Depends(get_session)):
    return session.exec(select(EngineIssue)).all()


# ========== 3️⃣ التأهيل ==========
@router.post("/rehab")
def add_engine_rehab(data: EngineRehab, session: Session = Depends(get_session)):
    data.date = data.date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/rehab")
def list_engine_rehab(session: Session = Depends(get_session)):
    return session.exec(select(EngineRehab)).all()


# ========== 4️⃣ الفحص ==========
@router.post("/check")
def add_engine_check(data: EngineCheck, session: Session = Depends(get_session)):
    data.date = data.date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/check")
def list_engine_check(session: Session = Depends(get_session)):
    return session.exec(select(EngineCheck)).all()


# ========== 5️⃣ الرفع ==========
@router.post("/upload")
def add_engine_upload(data: EngineUpload, session: Session = Depends(get_session)):
    data.rehab_uploaded = data.rehab_uploaded or "لا"
    data.check_uploaded = data.check_uploaded or "لا"
    data.rehab_date = data.rehab_date or datetime.now().strftime("%Y-%m-%d")
    data.check_date = data.check_date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/upload")
def list_engine_upload(session: Session = Depends(get_session)):
    return session.exec(select(EngineUpload)).all()


# ========== 6️⃣ المخرطة ==========
@router.post("/lathe")
def add_engine_lathe(data: EngineLathe, session: Session = Depends(get_session)):
    data.date = data.date or datetime.now().strftime("%Y-%m-%d")
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/lathe")
def list_engine_lathe(session: Session = Depends(get_session)):
    return session.exec(select(EngineLathe)).all()


# ========== 7️⃣ البمبات والنوزلات ==========
@router.post("/pump")
def add_engine_pump(data: EnginePump, session: Session = Depends(get_session)):
    data.rehab = data.rehab or ""
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/pump")
def list_engine_pump(session: Session = Depends(get_session)):
    return session.exec(select(EnginePump)).all()


# ========== 8️⃣ الصريمي ==========
@router.post("/electrical")
def add_engine_electrical(
    data: EngineElectrical, session: Session = Depends(get_session)
):
    data.start = data.start or "لا"
    data.dynamo = data.dynamo or "لا"
    session.add(data)
    session.commit()
    session.refresh(data)
    return data


@router.get("/electrical")
def list_engine_electrical(session: Session = Depends(get_session)):
    return session.exec(select(EngineElectrical)).all()
