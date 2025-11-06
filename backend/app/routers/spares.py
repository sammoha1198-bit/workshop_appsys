from fastapi import APIRouter, Depends
from sqlmodel import Session

from ..database import get_session
from ..models import SparePart

router = APIRouter(prefix="/spares", tags=["spares"])

@router.post("")
def create_spare(item: SparePart, session: Session = Depends(get_session)):
    session.add(item)
    session.commit()
    session.refresh(item)
    return item
