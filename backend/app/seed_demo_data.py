# app/seed_demo_data.py
from sqlmodel import Session, select
from .database import engine, init_db
from . import models


def run_seed() -> dict:
    """
    يضيف بيانات تجريبية فقط لو مافيش بيانات.
    ينفع نستدعيه من /api/seed في main.py
    """
    init_db()

    with Session(engine) as s:
        # لو في بيانات توريد محركات خلاص ما نعيد
        if s.exec(select(models.EngineSupply)).first():
            return {"ok": False, "error": "already-seeded"}

        # ================== محركات ==================
        engines_supply = [
            models.EngineSupply(
                serial="111",
                engineType="Deutz",
                model="F4L912",
                prevSite="المخزن الرئيسي",
                supDate="2025-10-30",
                supplier="Yemen Mobile",
                notes="توريد جديد",
            ),
            models.EngineSupply(
                serial="222",
                engineType="Perkins",
                model="404D-22",
                prevSite="فرع الحديدة",
                supDate="2025-10-29",
                supplier="PowerTech",
                notes="محرك جاهز للاستخدام",
            ),
            models.EngineSupply(
                serial="333",
                engineType="Kubota",
                model="V2203",
                prevSite="فرع صنعاء",
                supDate="2025-10-28",
                supplier="DieselPro",
                notes="محرك تم إعادة تأهيله",
            ),
        ]
        s.add_all(engines_supply)

        s.add_all([
            models.EngineIssue(
                serial="111",
                currSite="محطة تعز",
                receiver="المهندس سامي",
                requester="قسم التشغيل",
                issueDate="2025-11-01",
                notes="تم صرفه للموقع",
            ),
            models.EngineRehab(
                serial="333",
                rehabber="فريق التأهيل",
                rehabType="إصلاح كامل",
                rehabDate="2025-11-02",
                notes="تم تغيير حلقات ومضخة",
            ),
            models.EngineCheck(
                serial="222",
                inspector="فريق الفحص",
                description="فحص حرارة وضغط زيت",
                checkDate="2025-11-03",
                notes="الفحص ممتاز",
            ),
            models.EngineUpload(
                serial="111",
                rehabUp="نعم",
                checkUp="لا",
                rehabUpDate="2025-11-04",
                notes="تم رفع المؤهل فقط",
            ),
            models.EngineLathe(
                serial="333",
                lathe="تشغيل عمود + جلنبر",
                latheDate="2025-11-05",
                notes="مخرطة خارجية",
            ),
            models.EnginePump(
                serial="111",
                pumpSerial="P-111-A",
                pumpRehab="تنظيف ورش",
                notes="بمب جاهز",
            ),
            models.EngineElectrical(
                serial="222",
                etype="كهرباء كاملة",
                starter="نعم",
                alternator="نعم",
                edate="2025-11-06",
                notes="تم فحص الدينمو",
            ),
        ])

        # ================== مولدات ==================
        gens_supply = [
            models.GeneratorSupply(
                code="GEN001",
                gType="30kVA",
                model="FG Wilson",
                prevSite="المستودع المركزي",
                supDate="2025-10-30",
                supplier="Yemen Mobile",
                vendor="PowerMax",
                notes="مولد جديد",
            ),
            models.GeneratorSupply(
                code="GEN002",
                gType="20kVA",
                model="Perkins",
                prevSite="فرع إب",
                supDate="2025-10-29",
                supplier="Yemen Mobile",
                vendor="EnergyTech",
                notes="مولد مستخدم",
            ),
            models.GeneratorSupply(
                code="GEN003",
                gType="15kVA",
                model="Deutz",
                prevSite="فرع تعز",
                supDate="2025-10-27",
                supplier="Yemen Mobile",
                vendor="DieselPro",
                notes="مولد مؤهل",
            ),
        ]
        s.add_all(gens_supply)

        s.add_all([
            models.GeneratorIssue(
                code="GEN001",
                issueDate="2025-11-01",
                receiver="موقع ذمار",
                requester="قسم الطاقة",
                currSite="ذمار",
                notes="سليم",
            ),
            models.GeneratorInspect(
                code="GEN002",
                inspector="فريق الفحص",
                elecRehab="نعم",
                rehabDate="2025-11-02",
                rehabUp="نعم",
                checkUp="نعم",
                notes="تم الرفع للنظام",
            ),
        ])

        s.commit()

    return {"ok": True, "seeded": True}