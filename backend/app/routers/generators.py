from fastapi import APIRouter, Depends, HTTPException
from sqlmodel import Session, select
from typing import List
from datetime import date

from ..database import get_session
from ..models import GenSupply, GenIssue, GenInspect

router = APIRouter(prefix="/generators", tags=["generators"])

@router.post("/supply")
def create_gen_supply(item: GenSupply, session: Session = Depends(get_session)):
    if not item.code or not item.type:
        raise HTTPException(status_code=400, detail="code & type are required")
    if not item.date:
        item.date = date.today()
    session.add(item); session.commit(); session.refresh(item); return item

@router.post("/issue")
def create_gen_issue(item: GenIssue, session: Session = Depends(get_session)):
    if not item.code:
        raise HTTPException(status_code=400, detail="code is required")
    if not item.date:
        item.date = date.today()
    session.add(item); session.commit(); session.refresh(item); return item

@router.post("/inspect")
def create_gen_inspect(item: GenInspect, session: Session = Depends(get_session)):
    if not item.code:
        raise HTTPException(status_code=400, detail="code is required")
    if not item.date:
        item.date = date.today()
    session.add(item); session.commit(); session.refresh(item); return item

@router.get("/last3")
def last_three_generators(session: Session = Depends(get_session)):
    rows: List[GenSupply] = session.exec(
        select(GenSupply).order_by(GenSupply.date.desc(), GenSupply.id.desc())
    ).all()
    return [{"code": r.code, "prev_site": r.prev_site} for r in rows[:3]]
