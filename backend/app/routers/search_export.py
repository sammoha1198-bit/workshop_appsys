from typing import Dict, List, Any
from fastapi import APIRouter, Depends, Query
from fastapi.responses import StreamingResponse
from sqlmodel import Session, select
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

from ..database import get_session
from ..models import (
    EngineSupply, EngineIssue, EngineRehab, EngineCheck, EngineUpload, EngineLathe, EnginePump, EngineElectrical,
    GenSupply, GenIssue, GenInspect, SparePart
)


from fastapi.responses import StreamingResponse
from io import BytesIO
from openpyxl import Workbook
from ..utils import write_rows
from datetime import datetime

router = APIRouter(prefix="/search", tags=["Search & Export"])

def latest_one(session: Session, model, where):
    return session.exec(select(model).where(where).order_by(desc(model.created_at)).limit(1)).first()

@router.get("/summary")
def summary(q: str = Query(...), session: Session = Depends(get_session)) -> Dict[str, Any]:
    rows: List[Dict[str, Any]] = []

    # ========= محركات =========
    eng_sup = session.exec(select(EngineSupply).where(EngineSupply.serial == q)).all()
    for s in eng_sup:
        issue = latest_one(session, EngineIssue, EngineIssue.serial == s.serial)
        rehab = latest_one(session, EngineRehab, EngineRehab.serial == s.serial)
        check = latest_one(session, EngineCheck, EngineCheck.serial == s.serial)
        up    = latest_one(session, EngineUpload, EngineUpload.serial == s.serial)
        spare = latest_one(session, Spare, Spare.serial == s.serial)

        rows.append({
            "kind": "محرك",
            "key": s.serial,
            "supply_type": s.type,
            "supply_model": s.model,
            "supply_prev_site": s.prev_site or "",
            "supply_supplier": s.supplier or "",
            "supply_date": str(s.date or ""),
            "issue_current_site": issue.current_site if issue else "",
            "issue_date": str(issue.date) if issue and issue.date else "",
            "rehab_by": rehab.rehab_by if rehab else "",
            "rehab_type": rehab.rehab_type if rehab else "",
            "rehab_date": str(rehab.date) if rehab and rehab.date else "",
            "check_inspector": check.inspector if check else "",
            "check_date": str(check.date) if check and check.date else "",
            "upload_rehab_file": up.rehab_file if up else "",
            "upload_check_file": up.check_file if up else "",
            "spare_last": f"{spare.item} ×{spare.qty}" if spare else "",
        })

    # ========= مولدات =========
    gen_sup = session.exec(select(GenSupply).where(GenSupply.code == q)).all()
    for s in gen_sup:
        issue = latest_one(session, GenIssue, GenIssue.code == s.code)
        insp  = latest_one(session, GenInspect, GenInspect.code == s.code)
        spare = latest_one(session, Spare, Spare.serial == s.code)

        rows.append({
            "kind": "مولد",
            "key": s.code,
            "supply_type": s.type,
            "supply_model": s.model,
            "supply_prev_site": s.prev_site or "",
            "supply_supplier": s.supplier or "",
            "supply_entity": s.entity or "",
            "supply_date": str(s.date or ""),
            "issue_current_site": issue.current_site if issue else "",
            "issue_date": str(issue.date) if issue and issue.date else "",
            "rehab_by": insp.rehab if insp else "",
            "rehab_type": "",
            "rehab_date": str(insp.date) if insp and insp.date else "",
            "check_inspector": insp.inspector if insp else "",
            "check_date": str(insp.date) if insp and insp.date else "",
            "upload_rehab_file": insp.rehab_file if insp else "",
            "upload_check_file": insp.check_file if insp else "",
            "spare_last": f"{spare.item} ×{spare.qty}" if spare else "",
        })

    return {"rows": rows}

@router.get("/export")
def export_search(q: str = Query(...), session: Session = Depends(get_session)):
    data = summary(q, session)
    rows = data.get("rows", [])
    headers = [
        "النوع","الرقم","نوع/مودل","الموقع السابق","الموقع الحالي","المؤهل","الفاحص",
        "رفع مؤهل/فحص","آخر قطع","تاريخ التوريد","تاريخ الصرف","تاريخ التأهيل","تاريخ الفحص"
    ]
    xrows: List[List[str]] = []
    for r in rows:
        xrows.append([
            r.get("kind",""),
            r.get("key",""),
            f"{r.get('supply_type','')}/{r.get('supply_model','')}".strip("/"),
            r.get("supply_prev_site",""),
            r.get("issue_current_site",""),
            r.get("rehab_by","") or r.get("rehab_type",""),
            r.get("check_inspector",""),
            f"{r.get('upload_rehab_file','')}/{r.get('upload_check_file','')}".strip("/"),
            r.get("spare_last",""),
            r.get("supply_date",""),
            r.get("issue_date",""),
            r.get("rehab_date",""),
            r.get("check_date",""),
        ])

    wb = Workbook(); ws = wb.active; ws.title = "نتيجة البحث"
    write_rows(ws, headers, xrows)
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    fname = f"Search_{q}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename=\"{fname}\"'}
    )
