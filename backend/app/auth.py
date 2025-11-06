# C:\Workshop-System\backend\app\auth.py
from datetime import datetime, timedelta
from typing import Optional, List
import os, jwt
from fastapi import APIRouter, HTTPException, Depends, Header, Body, Query
from passlib.context import CryptContext
from sqlmodel import SQLModel, Session, select
from .database import get_session
from .models_user import User

JWT_SECRET = os.getenv("JWT_SECRET", "change-me")
JWT_EXPIRE_MINUTES = 24*60
pwd = CryptContext(schemes=["bcrypt"], deprecated="auto")
router = APIRouter(prefix="/auth", tags=["auth"])

def hash_pw(p): return pwd.hash(p)
def verify_pw(p, h): return pwd.verify(p, h)
def create_token(uid: int, role: str, name: str):
    payload = {"uid": uid, "role": role, "name": name, "exp": datetime.utcnow()+timedelta(minutes=JWT_EXPIRE_MINUTES)}
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")

class RegisterBody(SQLModel, table=False):
    email: Optional[str] = None
    name: Optional[str] = None
    password: Optional[str] = None
    role: Optional[str] = "viewer"
    admin_secret: Optional[str] = None

class LoginBody(SQLModel, table=False):
    email: Optional[str] = None
    password: Optional[str] = None

def is_empty_users(session: Session) -> bool:
    return session.exec(select(User).limit(1)).first() is None

@router.post("/register")
def register(
    # Query (تعمل لوحدها)
    email_q: Optional[str] = Query(None),
    name_q: Optional[str] = Query(None),
    password_q: Optional[str] = Query(None),
    role_q: Optional[str] = Query("viewer"),
    admin_secret_q: Optional[str] = Query(None),
    # JSON (بديل)
    body: Optional[RegisterBody] = Body(None),
    session: Session = Depends(get_session),
):
    email = email_q or (body.email if body else None)
    name = name_q or (body.name if body else None)
    password = password_q or (body.password if body else None)
    role = role_q or (body.role if body else "viewer")
    admin_secret = admin_secret_q or (body.admin_secret if body else None)

    if not email or not password or not name:
        raise HTTPException(status_code=422, detail="missing fields")

    # تمهيد أوّل مشرف تلقائيًا إذا لا يوجد أي مستخدم
    if is_empty_users(session):
        # إن لم يحدد دورًا، نجعله admin افتراضيًا في التمهيد
        if role.lower() != "admin":
            role = "admin"
        u = User(email=email, name=name, role=role, password_hash=hash_pw(password))
        session.add(u); session.commit(); session.refresh(u)
        return {"ok": True, "bootstrap": True, "id": u.id}

    # بعد وجود مستخدم واحد على الأقل → يجب مطابقة السر
    MASTER = os.getenv("ADMIN_CREATE_SECRET", "")
    if not admin_secret or admin_secret != MASTER:
        raise HTTPException(status_code=403, detail="not allowed")

    # منع تكرار نفس الايميل
    if session.exec(select(User).where(User.email == email)).first():
        return {"ok": True, "exists": True}

    u = User(email=email, name=name, role=role, password_hash=hash_pw(password))
    session.add(u); session.commit(); session.refresh(u)
    return {"ok": True, "id": u.id}

@router.post("/login")
def login(
    email_q: Optional[str] = Query(None),
    password_q: Optional[str] = Query(None),
    body: Optional[LoginBody] = Body(None),
    session: Session = Depends(get_session),
):
    email = email_q or (body.email if body else None)
    password = password_q or (body.password if body else None)
    if not email or not password:
        raise HTTPException(status_code=422, detail="missing fields")

    u = session.exec(select(User).where(User.email == email)).first()
    if not u or not verify_pw(password, u.password_hash):
        raise HTTPException(status_code=401, detail="bad creds")

    return {"token": create_token(u.id, u.role, u.name), "role": u.role, "name": u.name}

def get_user(authorization: Optional[str] = Header(None)):
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="no token")
    token = authorization.split(" ",1)[1]
    try:
        data = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
        return data
    except Exception:
        raise HTTPException(status_code=401, detail="bad token")

def role_required(roles: List[str]):
    def dep(user=Depends(get_user)):
        if user["role"] not in roles:
            raise HTTPException(status_code=403, detail="forbidden")
        return user
    return dep
