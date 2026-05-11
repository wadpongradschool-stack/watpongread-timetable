"""
ระบบจัดตารางเรียนตารางสอนอัตโนมัติ
Streamlit Version สำหรับ Deploy บน GitHub / Streamlit Cloud
"""

from copy import deepcopy
from io import BytesIO
import random

# =========================
# SAFE IMPORTS
# =========================

import streamlit as st


try:
    import pandas as pd
except ModuleNotFoundError:
    raise ModuleNotFoundError(
        "กรุณาติดตั้ง pandas ด้วยคำสั่ง: pip install pandas"
    )

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError:
    raise ModuleNotFoundError(
        "กรุณาติดตั้ง openpyxl ด้วยคำสั่ง: pip install openpyxl"
    )

# =========================
# CONSTANTS
# =========================

DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"]
PERIODS = [1, 2, 3, 4, 5, 6]
MORNING = [1, 2, 3]

DEFAULT_TEACHERS = [
    {
        "id": 1,
        "name": "ครู ณัฐวุฒิ",
        "specialty": "สุขศึกษาและพลศึกษา",
        "homeroom": "ป.5",
    },
    {
        "id": 2,
        "name": "ครู ชมพูนุท",
        "specialty": "อังกฤษพื้นฐานและอังกฤษเพิ่มเติม",
        "homeroom": "ป.6",
    },
    {
        "id": 3,
        "name": "ครู มณีมัญช์",
        "specialty": "วิทยาศาสตร์",
        "homeroom": "ป.4",
    },
    {
        "id": 4,
        "name": "ครู กชกร",
        "specialty": "ภาษาไทย",
        "homeroom": "ป.2",
    },
    {
        "id": 5,
        "name": "ครู รักชนก",
        "specialty": "วิทยาการคำนวณ",
        "homeroom": "ป.3",
    },
    {
        "id": 6,
        "name": "ครู ณิชนันทน์",
        "specialty": "สอนทุกวิชา",
        "homeroom": "ป.1",
    },
    {
        "id": 7,
        "name": "ครู ณัฐนนท์",
        "specialty": "ดนตรี",
        "homeroom": "-",
    },
]

DEFAULT_CLASSES = ["ป.1", "ป.2", "ป.3", "ป.4", "ป.5", "ป.6"]

DEFAULT_SUBJECTS = [
    {"id": 1, "name": "ภาษาไทย", "periods": 5, "needs_focus": True},
    {"id": 2, "name": "คณิตศาสตร์", "periods": 5, "needs_focus": True},
    {"id": 3, "name": "วิทยาศาสตร์", "periods": 3, "needs_focus": True},
    {"id": 4, "name": "สังคมศึกษาฯ", "periods": 2, "needs_focus": False},
    {"id": 5, "name": "ประวัติศาสตร์", "periods": 1, "needs_focus": False},
    {"id": 6, "name": "สุขศึกษา", "periods": 1, "needs_focus": False},
    {"id": 7, "name": "ศิลปะ", "periods": 2, "needs_focus": False},
    {"id": 8, "name": "การงานอาชีพ", "periods": 2, "needs_focus": False},
    {"id": 9, "name": "อังกฤษ", "periods": 4, "needs_focus": True},
    {"id": 10, "name": "อังกฤษ(เพิ่มเติม)", "periods": 1, "needs_focus": True},
    {"id": 11, "name": "ป้องกันการทุจริต", "periods": 1, "needs_focus": False},
    {"id": 12, "name": "แนะแนว", "periods": 1, "needs_focus": False},
    {"id": 13, "name": "ลูกเสือ เนตรนารี", "periods": 1, "needs_focus": False},
    {"id": 14, "name": "ชุมนุม", "periods": 1, "needs_focus": False},
    {"id": 15, "name": "พลศึกษา", "periods": 2, "needs_focus": False},
    {"id": 16, "name": "สวดมนต์", "periods": 1, "needs_focus": False},
    {"id": 17, "name": "วิทยาการคำนวณ", "periods": 2, "needs_focus": True},
    {"id": 18, "name": "ดนตรี", "periods": 1, "needs_focus": False},
]

DEFAULT_TEACHER_SUBJECTS = {
    1: ["สุขศึกษา", "พลศึกษา", "ลูกเสือ เนตรนารี"],
    2: ["อังกฤษ", "อังกฤษ(เพิ่มเติม)"],
    3: ["วิทยาศาสตร์"],
    4: ["ภาษาไทย"],
    5: ["วิทยาการคำนวณ"],
    6: [
        "ภาษาไทย",
        "คณิตศาสตร์",
        "วิทยาศาสตร์",
        "สังคมศึกษาฯ",
        "ประวัติศาสตร์",
        "สุขศึกษา",
        "ศิลปะ",
        "การงานอาชีพ",
        "อังกฤษ",
        "อังกฤษ(เพิ่มเติม)",
        "ป้องกันการทุจริต",
        "แนะแนว",
        "ลูกเสือ เนตรนารี",
        "ชุมนุม",
        "พลศึกษา",
        "สวดมนต์",
        "วิทยาการคำนวณ",
    ],
    7: ["ดนตรี"],
}

MUSIC_SCHEDULE = {
    "จันทร์": ["ป.1", "ป.2"],
    "อังคาร": [],
    "พุธ": ["ป.3", "ป.4"],
    "พฤหัสบดี": ["ป.5", "ป.6"],
    "ศุกร์": ["กิจกรรมกลองยาว"],
}

MUSIC_PERIOD = 4  # 12:30 - 13:30

# =========================
# CORE FUNCTIONS
# =========================


def generate_timetable(classes, subjects, teachers, teacher_subjects):
    """Generate timetable safely without Streamlit dependency."""

    random.seed(42)

    schedule = {
        cls: {
            day: {period: None for period in PERIODS}
            for day in DAYS
        }
        for cls in classes
    }

    teacher_busy = {
        teacher["id"]: {
            day: {period: False for period in PERIODS}
            for day in DAYS
        }
        for teacher in teachers
    }

    conflicts = []

    tasks = []

    for cls in classes:
        for subject in subjects:
            subject_name = subject["name"]

            eligible_teachers = [
                teacher["id"]
                for teacher in teachers
                if subject_name in teacher_subjects.get(teacher["id"], [])
            ]

            if not eligible_teachers:
                conflicts.append(f"ไม่มีครูสอนวิชา {subject_name}")
                continue

            teacher_id = eligible_teachers[0]

            for _ in range(subject["periods"]):
                tasks.append({
                    "class": cls,
                    "subject": subject,
                    "teacher_id": teacher_id,
                })

    tasks.sort(
        key=lambda x: (
            0 if x["subject"].get("needs_focus") else 1,
            -x["subject"]["periods"],
        )
    )

    for cls in classes:
        # ล็อกคาบพิเศษ (กิจกรรมรวมครูทุกคน)
        wed_teacher = 0
        fri_teacher = 0

        # พุธคาบสุดท้าย = ลูกเสือ (ครูทุกคน)
        schedule[cls]["พุธ"][6] = {
            "subject": "ลูกเสือ เนตรนารี",
            "teacher_id": wed_teacher,
        }

        # ศุกร์คาบสุดท้าย = สวดมนต์ (ครูทุกคน)
        schedule[cls]["ศุกร์"][6] = {
            "subject": "สวดมนต์",
            "teacher_id": fri_teacher,
        }

    for task in tasks:
        cls = task["class"]
        subject = task["subject"]
        teacher_id = task["teacher_id"]

        # ข้ามวิชาที่ล็อกไว้แล้ว
        if subject["name"] in ["ลูกเสือ เนตรนารี", "สวดมนต์"]:
            continue

        placed = False

        for day in DAYS:
            for period in PERIODS:

                # ข้ามคาบล็อกพิเศษ
                if (day == "พุธ" and period == 6) or (
                    day == "ศุกร์" and period == 6
                ):
                    continue

                if schedule[cls][day][period] is not None:
                    continue

                if teacher_busy[teacher_id][day][period]:
                    continue

                schedule[cls][day][period] = {
                    "subject": subject["name"],
                    "teacher_id": teacher_id,
                }

                teacher_busy[teacher_id][day][period] = True
                placed = True
                break

            if placed:
                break

        if not placed:
            conflicts.append(
                f"ไม่สามารถจัด {subject['name']} ให้ {cls} ได้"
            )

    return schedule, conflicts


def calculate_teacher_workload(schedule, teachers):
    """คำนวณภาระงานครู"""

    workload = []

    for teacher in teachers:
        total_periods = 0

        for cls in schedule:
            for day in DAYS:
                for period in PERIODS:
                    slot = schedule[cls][day][period]

                    if slot and slot["teacher_id"] == teacher["id"]:
                        total_periods += 1

        if total_periods >= 25:
            level = "มาก"
        elif total_periods >= 15:
            level = "ปานกลาง"
        else:
            level = "น้อย"

        workload.append({
            "ครู": teacher["name"],
            "ประจำชั้น": teacher["homeroom"],
            "วิชา": teacher["specialty"],
            "จำนวนคาบ": total_periods,
            "ภาระงาน": level,
        })

    return pd.DataFrame(workload)


# =========================
# EXCEL EXPORT
# =========================


def export_excel(schedule, filename="timetable.xlsx"):
    """Export schedule to Excel file."""

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "ตารางเรียน"

    headers = ["ชั้นเรียน", "วัน", "คาบ", "วิชา", "ครู"]

    for col, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col, value=header)

    row = 2

    for cls, class_schedule in schedule.items():
        for day, periods in class_schedule.items():
            for period, slot in periods.items():
                if slot:
                    sheet.cell(row=row, column=1, value=cls)
                    sheet.cell(row=row, column=2, value=day)
                    sheet.cell(row=row, column=3, value=period)
                    sheet.cell(row=row, column=4, value=slot["subject"])
                    sheet.cell(row=row, column=5, value=slot["teacher_id"])
                    row += 1

    workbook.save(filename)


# =========================
# TESTS
# =========================


def run_basic_tests():
    """Simple internal tests."""

    schedule, conflicts = generate_timetable(
        classes=deepcopy(DEFAULT_CLASSES),
        subjects=deepcopy(DEFAULT_SUBJECTS),
        teachers=deepcopy(DEFAULT_TEACHERS),
        teacher_subjects=deepcopy(DEFAULT_TEACHER_SUBJECTS),
    )

    assert isinstance(schedule, dict), "schedule ต้องเป็น dict"
    assert isinstance(conflicts, list), "conflicts ต้องเป็น list"

    total_slots = 0

    for cls in schedule:
        for day in DAYS:
            for period in PERIODS:
                if schedule[cls][day][period]:
                    total_slots += 1

    assert total_slots > 0, "ต้องมีวิชาถูกจัดอย่างน้อย 1 คาบ"

    print("✅ All tests passed")


# =========================
# MAIN
# =========================

st.set_page_config(
    page_title="ระบบจัดตารางเรียน",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

.main-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%);
    padding: 2rem;
    border-radius: 22px;
    color: white;
    margin-bottom: 1.5rem;
    box-shadow: 0 10px 30px rgba(37,99,235,0.25);
}

.main-header h1 {
    font-size: 2.2rem;
    font-weight: 800;
    margin-bottom: 0.2rem;
}

.main-header p {
    opacity: 0.9;
    font-size: 1rem;
}

.metric-card {
    background: white;
    padding: 1rem;
    border-radius: 18px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.08);
    border: 1px solid #e2e8f0;
}

.stButton>button {
    width: 100%;
    border-radius: 12px;
    border: none;
    background: linear-gradient(135deg,#2563eb,#1d4ed8);
    color: white;
    font-weight: 700;
    padding: 0.65rem 1rem;
    transition: all 0.2s ease;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 18px rgba(37,99,235,0.3);
}

.stTabs [data-baseweb="tab"] {
    font-weight: 700;
    border-radius: 10px;
}

.stDataFrame {
    border-radius: 16px;
    overflow: hidden;
}

section[data-testid="stSidebar"] {
    background: #f8fafc;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
    <h1>🏫 ระบบจัดตารางเรียนอัตโนมัติ</h1>
    <p>โรงเรียนวัดโป่งแรด • Smart School Timetable System</p>
</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(f"""
    <div class='metric-card'>
        <h3>👩‍🏫 ครูทั้งหมด</h3>
        <h1>{len(DEFAULT_TEACHERS)}</h1>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class='metric-card'>
        <h3>📚 รายวิชา</h3>
        <h1>{len(DEFAULT_SUBJECTS)}</h1>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class='metric-card'>
        <h3>🏫 ระดับชั้น</h3>
        <h1>{len(DEFAULT_CLASSES)}</h1>
    </div>
    """, unsafe_allow_html=True)

st.sidebar.markdown("## ⚙️ แผงควบคุม")
st.sidebar.markdown("จัดการและสร้างตารางเรียนอัตโนมัติ")

if st.sidebar.button("🚀 สร้างตารางเรียน"):

    schedule_result, conflict_result = generate_timetable(
        classes=deepcopy(DEFAULT_CLASSES),
        subjects=deepcopy(DEFAULT_SUBJECTS),
        teachers=deepcopy(DEFAULT_TEACHERS),
        teacher_subjects=deepcopy(DEFAULT_TEACHER_SUBJECTS),
    )

    st.session_state["schedule"] = schedule_result
    st.session_state["conflicts"] = conflict_result

    st.success("สร้างตารางสำเร็จ")

if "schedule" in st.session_state:

    schedule = st.session_state["schedule"]
    conflicts = st.session_state["conflicts"]

    if conflicts:
        st.warning("พบ Conflicts")
        for item in conflicts:
            st.write(f"- {item}")
    else:
        st.success("✅ ไม่มี Conflict")

    tabs = st.tabs(DEFAULT_CLASSES)

    for idx, cls in enumerate(DEFAULT_CLASSES):

        with tabs[idx]:

            rows = []

            for period in PERIODS:
                row = {"คาบ": f"คาบ {period}"}

                for day in DAYS:
                    slot = schedule[cls][day][period]

                    if slot:
                        teacher_name = (
                            "ครูทุกคน"
                            if slot["teacher_id"] == 0
                            else next(
                                (
                                    teacher["name"]
                                    for teacher in DEFAULT_TEACHERS
                                    if teacher["id"] == slot["teacher_id"]
                                ),
                                "-",
                            )
                        )

                        row[day] = f"{slot['subject']} ({teacher_name})"
                    else:
                        row[day] = "-"

                rows.append(row)

            st.dataframe(pd.DataFrame(rows), use_container_width=True)

    export_excel(schedule)

    with open("timetable.xlsx", "rb") as file:
        st.download_button(
            label="📥 ดาวน์โหลด Excel",
            data=file,
            file_name="timetable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    st.subheader("📊 สรุปภาระงานครู")

    workload_df = calculate_teacher_workload(
        schedule,
        DEFAULT_TEACHERS,
    )

    st.dataframe(workload_df, use_container_width=True)

    high_load = workload_df[
        workload_df["ภาระงาน"] == "มาก"
    ]

    low_load = workload_df[
        workload_df["ภาระงาน"] == "น้อย"
    ]

    col1, col2 = st.columns(2)

    with col1:
        st.metric("👨‍🏫 ภาระงานมาก", len(high_load))

    with col2:
        st.metric("📘 ภาระงานน้อย", len(low_load))

st.markdown("---")

st.markdown("---")
st.subheader("👩‍🏫 ข้อมูลครูผู้สอน")
st.dataframe(pd.DataFrame(DEFAULT_TEACHERS), use_container_width=True)

st.subheader("📚 ข้อมูลรายวิชา")
st.dataframe(pd.DataFrame(DEFAULT_SUBJECTS), use_container_width=True)

run_basic_tests()
