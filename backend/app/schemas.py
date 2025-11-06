from pydantic import BaseModel
from datetime import date
from typing import Optional

class DateRange(BaseModel):
    start: Optional[date]
    end: Optional[date]
