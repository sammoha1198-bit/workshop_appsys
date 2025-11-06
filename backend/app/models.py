from sqlmodel import SQLModel, Field
from typing import Optional
import datetime as dt

def utcnow() -> str:
    return dt.datetime.utcnow().isoformat()

# ------------ محركات ------------
class EngineSupply(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    engineType: Optional[str] = None
    model: Optional[str] = None
    prevSite: Optional[str] = None
    supDate: Optional[str] = None
    supplier: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineIssue(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    currSite: Optional[str] = None
    receiver: Optional[str] = None
    requester: Optional[str] = None
    issueDate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineRehab(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    rehabber: Optional[str] = None
    rehabType: Optional[str] = None
    rehabDate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineCheck(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    inspector: Optional[str] = None
    desc: Optional[str] = None   # (مهم) بعض النسخ القديمة لا تحتويه
    checkDate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineUpload(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    rehabUp: Optional[str] = None
    checkUp: Optional[str] = None
    rehabUpDate: Optional[str] = None
    checkUpDate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineLathe(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    lathe: Optional[str] = None
    latheDate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EnginePump(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    pumpSerial: Optional[str] = None
    pumpRehab: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class EngineElectrical(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    serial: str = Field(index=True)
    etype: Optional[str] = None
    starter: Optional[str] = None
    alternator: Optional[str] = None
    edate: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

# ------------ مولدات ------------
class GeneratorSupply(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True)
    gType: Optional[str] = None
    model: Optional[str] = None
    prevSite: Optional[str] = None
    supDate: Optional[str] = None
    supplier: Optional[str] = None
    vendor: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class GeneratorIssue(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True)
    issueDate: Optional[str] = None
    receiver: Optional[str] = None
    requester: Optional[str] = None
    currSite: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)

class GeneratorInspect(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True)
    inspector: Optional[str] = None
    elecRehab: Optional[str] = None
    rehabDate: Optional[str] = None
    rehabUp: Optional[str] = None
    checkUp: Optional[str] = None
    notes: Optional[str] = None
    created_at: str = Field(default_factory=utcnow)
