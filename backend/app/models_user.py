# C:\Workshop-System\backend\app\models_user.py
from typing import Optional
from sqlmodel import SQLModel, Field

class User(SQLModel, table=True):
    __tablename__ = "users"
    id: Optional[int] = Field(default=None, primary_key=True)
    email: str = Field(index=True)
    name: str
    role: str = Field(default="viewer")  # admin | tech | viewer
    password_hash: str
