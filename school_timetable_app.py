"""
ระบบจัดตารางเรียนตารางสอนอัตโนมัติ
โรงเรียนวัดโป่งแรด
"""

import streamlit as st
import pandas as pd
import random
from copy import deepcopy
from io import BytesIO
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ─── Page Config ────────────────────────────────────────────
st.set_page_config(
    page_title="ตารางสอน | วัดโป่งแรด",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

/* Header */
.main-header {
    background: linear-gradient(135deg, #1a472a 0%, #2d6a4f 50%, #40916c 100%);
    padding: 2rem 2.5rem;
    border-radius: 16px;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(26,71,42,0.3);
}
.main-header h1 {
    color: #ffffff;
    font-size: 2rem;
    font-weight: 700;
    margin: 0;
}
.main-header p {
    color: #b7e4c7;
    font-size: 1rem;
    margin: 0.3rem 0 0 0;
}

/* Stats cards */
.stat-card {
    background: white;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    border-left: 5px solid #40916c;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    margin-bottom: 1rem;
}
.stat-number {
    font-size: 2.2rem;
    font-weight: 700;
    color: #1a472a;
    line-height: 1;
}
.stat-label {
    font-size: 0.9rem;
    color: #6c757d;
    margin-top: 0.3rem;
}

/* Section headers */
.section-title {
    font-size: 1.2rem;
    font-weight: 700;
    color: #1a472a;
    border-bottom: 3px solid #40916c;
    padding-bottom: 0.5rem;
    margin: 1.5rem 0 1rem 0;
}

/* Timetable cells */
.period-morning {
    background-color: #d8f3dc;
    color: #1a472a;
    border-radius: 6px;
    padding: 4px 8px;
    font-size: 0.8rem;
    font-weight: 600;
}
.period-afternoon {
    background-color: #fff3cd;
    color: #856404;
    border-radius: 6px;
    padding: 4px 8px;
    font-size: 0.8rem;
    font-weight: 600;
}
.period-free {
    color: #adb5bd;
    font-size: 0.8rem;
}

/* Conflict badge */
.conflict-badge {
    background: #f8d7da;
    color: #842029;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.8rem;
    font-weight: 600;
}
.ok-badge {
    background: #d1e7dd;
    color: #0f5132;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.8rem;
    font-weight: 600;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #1a472a, #40916c);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 1.5rem;
    font-family: 'Sarabun', sans-serif;
    font-weight: 600;
    font-size: 1rem;
    transition: all 0.2s;
    box-shadow: 0 4px 12px rgba(26,71,42,0.25);
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(26,71,42,0.4);
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background: #f8f9fa;
    border-radius: 10px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    font-family: 'Sarabun', sans-serif;
    font-weight: 600;
}
.stTabs [aria-selected="true"] {
    background: #1a472a !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)


# ─── Default Data ────────────────────────────────────────────
DEFAULT_TEACHERS = [
    {"id": 1,  "name": "นายสมศักดิ์ ใจดี",    "specialty": "ภาษาไทย"},
    {"id": 2,  "name": "นางวิไล มีสุข",        "specialty": "คณิตศาสตร์"},
    {"id": 3,  "name": "นายประสิทธิ์ เก่งงาน", "specialty": "วิทยาศาสตร์"},
    {"id": 4,  "name": "นางสาวสุดา รักเรียน",  "specialty": "ภาษาอังกฤษ"},
    {"id": 5,  "name": "นายชาญชัย มั่นคง",     "specialty": "สังคมศึกษา"},
    {"id": 6,  "name": "นางนิตยา สวยงาม",      "specialty": "ศิลปะ"},
    {"id": 7,  "name": "นายวีระ แข็งแรง",      "specialty": "พลศึกษา"},
    {"id": 8,  "name": "นางสาวปิยะ ขยันดี",    "specialty": "การงานอาชีพ"},
    {"id": 9,  "name": "นายอนันต์ รู้รอบ",     "specialty": "คณิตศาสตร์"},
    {"id": 10, "name": "นางมาลี ตั้งใจ",       "specialty": "ภาษาไทย"},
    {"id": 11, "name": "นายธีรพล ฉลาด",        "specialty": "วิทยาศาสตร์"},
    {"id": 12, "name": "นางสาวรัตนา เปรียว",   "specialty": "ภาษาอังกฤษ"},
    {"id": 13, "name": "นายกิตติ เป็นเลิศ",    "specialty": "ดนตรี"},
    {"id": 14, "name": "นางพรทิพย์ ดูแลดี",    "specialty": "แนะแนว"},
]

DEFAULT_SUBJECTS = [
    {"id": 1, "name": "ภาษาไทย",     "code": "T101", "periods": 5, "needs_focus": True,  "room": "ห้องเรียน"},
    {"id": 2, "name": "คณิตศาสตร์",  "code": "M101", "periods": 5, "needs_focus": True,  "room": "ห้องเรียน"},
    {"id": 3, "name": "วิทยาศาสตร์", "code": "S101", "periods": 3, "needs_focus": True,  "room": "ห้องเรียน"},
    {"id": 4, "name": "ภาษาอังกฤษ",  "code": "E101", "periods": 4, "needs_focus": True,  "room": "ห้องเรียน"},
    {"id": 5, "name": "สังคมศึกษา",  "code": "SS101","periods": 3, "needs_focus": False, "room": "ห้องเรียน"},
    {"id": 6, "name": "ศิลปะ",       "code": "A101", "periods": 2, "needs_focus": False, "room": "ห้องเรียน"},
    {"id": 7, "name": "พลศึกษา",     "code": "PE101","periods": 2, "needs_focus": False, "room": "สนามกีฬา"},
    {"id": 8, "name": "การงานอาชีพ", "code": "V101", "periods": 2, "needs_focus": False, "room": "ห้องคอมพิวเตอร์"},
    {"id": 9, "name": "ดนตรี",       "code": "MU101","periods": 1, "needs_focus": False, "room": "ห้องดนตรี"},
    {"id": 10,"name": "แนะแนว",      "code": "G101", "periods": 1, "needs_focus": False, "room": "ห้องเรียน"},
]

DEFAULT_CLASSES = ["ป.1", "ป.2", "ป.3", "ป.4", "ป.5", "ป.6"]

DAYS     = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"]
PERIODS  = list(range(1, 9))   # 8 คาบ/วัน
MORNING  = [1, 2, 3, 4]

# Subject → Teacher mapping (simplified for demo)
SUBJECT_TEACHER = {
    "ภาษาไทย":     [1, 10],
    "คณิตศาสตร์":  [2, 9],
    "วิทยาศาสตร์": [3, 11],
    "ภาษาอังกฤษ":  [4, 12],
    "สังคมศึกษา":  [5],
    "ศิลปะ":       [6],
    "พลศึกษา":     [7],
    "การงานอาชีพ": [8],
    "ดนตรี":       [13],
    "แนะแนว":      [14],
}

SUBJECT_COLOR = {
    "ภาษาไทย":     "#d8f3dc",
    "คณิตศาสตร์":  "#caf0f8",
    "วิทยาศาสตร์": "#e2d9f3",
    "ภาษาอังกฤษ":  "#ffd6a5",
    "สังคมศึกษา":  "#ffccd5",
    "ศิลปะ":       "#f8edeb",
    "พลศึกษา":     "#d4edda",
    "การงานอาชีพ": "#fff3cd",
    "ดนตรี":       "#e8d5b7",
    "แนะแนว":      "#e2e3e5",
    "—":           "#f8f9fa",
}


# ─── Scheduler Logic ────────────────────────────────────────
def generate_timetable(classes, subjects, teachers_list, soft_focus_morning, soft_teacher_gap):
    """
    Simple greedy + constraint-check scheduler.
    Returns: schedule dict {class_name: {day: {period: {subject, teacher_id}}}}
    """
    random.seed(42)

    # Build teacher lookup by id
    teacher_by_id = {t["id"]: t for t in teachers_list}

    # Initialize empty schedule
    schedule = {}
    for cls in classes:
        schedule[cls] = {}
        for day in DAYS:
            schedule[cls][day] = {p: None for p in PERIODS}

    # Track teacher occupation: teacher_busy[teacher_id][day][period] = bool
    teacher_busy = {t["id"]: {d: {p: False for p in PERIODS} for d in DAYS}
                    for t in teachers_list}

    # Build task list: (class, subject, teacher_id) × needed periods
    tasks = []
    for cls in classes:
        for subj in subjects:
            sname   = subj["name"]
            needed  = subj["periods"]
            t_ids   = SUBJECT_TEACHER.get(sname, [1])
            # Distribute teachers across classes
            t_id = t_ids[classes.index(cls) % len(t_ids)]
            for _ in range(needed):
                tasks.append({
                    "class":      cls,
                    "subject":    subj,
                    "teacher_id": t_id,
                })

    # Sort: focus subjects first so they land in morning slots
    tasks.sort(key=lambda x: (
        0 if x["subject"]["needs_focus"] else 1,
        -x["subject"]["periods"]
    ))

    conflicts = []

    for task in tasks:
        cls     = task["class"]
        subj    = task["subject"]
        t_id    = task["teacher_id"]
        placed  = False

        # Build candidate slots, scored by soft constraints
        candidates = []
        for day in DAYS:
            for period in PERIODS:
                # Hard: class slot free?
                if schedule[cls][day][period] is not None:
                    continue
                # Hard: teacher free?
                if teacher_busy[t_id][day][period]:
                    continue

                score = 0
                # Soft: focus subject in morning
                if soft_focus_morning and subj["needs_focus"] and period in MORNING:
                    score += 10
                # Soft: teacher has at least 1 free period this day
                if soft_teacher_gap:
                    used = sum(1 for p in PERIODS if teacher_busy[t_id][day][p])
                    if used < len(PERIODS) - 1:
                        score += 3

                candidates.append((score, day, period))

        candidates.sort(key=lambda x: -x[0])

        if candidates:
            _, best_day, best_period = candidates[0]
            schedule[cls][best_day][best_period] = {
                "subject":    subj["name"],
                "teacher_id": t_id,
                "room":       subj["room"],
            }
            teacher_busy[t_id][best_day][best_period] = True
            placed = True

        if not placed:
            conflicts.append(f"{cls} — {subj['name']} (ครู ID {t_id})")

    return schedule, teacher_busy, conflicts, teacher_by_id


def build_teacher_schedule(teacher_id, schedule, classes):
    """Extract one teacher's schedule from full schedule."""
    result = {d: {p: None for p in PERIODS} for d in DAYS}
    for cls in classes:
        for day in DAYS:
            for period in PERIODS:
                slot = schedule[cls][day][period]
                if slot and slot["teacher_id"] == teacher_id:
                    result[day][period] = {"class": cls, "subject": slot["subject"]}
    return result


# ─── Excel Export ────────────────────────────────────────────
def export_to_excel(schedule, teacher_schedules, teachers_list, classes):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill   = PatternFill("solid", fgColor="1A472A")
    subhead_fill  = PatternFill("solid", fgColor="40916C")
    morning_fill  = PatternFill("solid", fgColor="D8F3DC")
    afternoon_fill= PatternFill("solid", fgColor="FFF3CD")
    free_fill     = PatternFill("solid", fgColor="F8F9FA")
    header_font   = Font(name="TH Sarabun New", bold=True, color="FFFFFF", size=13)
    cell_font     = Font(name="TH Sarabun New", size=11)
    bold_font     = Font(name="TH Sarabun New", bold=True, size=11)
    center        = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border   = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    def style_cell(cell, fill=None, font=None, align=center, border=thin_border):
        if fill:   cell.fill      = fill
        if font:   cell.font      = font
        if align:  cell.alignment = align
        if border: cell.border    = border

    # ── Sheet: ตารางเรียนแต่ละชั้น ──────────────────────────
    for cls in classes:
        ws = wb.create_sheet(title=f"ตาราง {cls}")
        ws.sheet_view.showGridLines = False

        # Title
        ws.merge_cells("A1:I1")
        ws["A1"] = f"ตารางเรียน ชั้น {cls}  —  โรงเรียนวัดโป่งแรด"
        ws["A1"].fill      = header_fill
        ws["A1"].font      = Font(name="TH Sarabun New", bold=True, color="FFFFFF", size=15)
        ws["A1"].alignment = center
        ws.row_dimensions[1].height = 36

        # Column headers
        ws.column_dimensions["A"].width = 6
        headers = ["คาบ"] + DAYS
        for col_i, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_i, value=h)
            cell.fill = subhead_fill
            style_cell(cell, font=header_font)
            ws.column_dimensions[get_column_letter(col_i)].width = 22

        # Rows
        for p_i, period in enumerate(PERIODS, 3):
            ws.cell(row=p_i, column=1, value=f"คาบ {period}")
            style_cell(ws.cell(row=p_i, column=1),
                       fill=PatternFill("solid", fgColor="E9ECEF"),
                       font=bold_font)
            ws.row_dimensions[p_i].height = 38

            for d_i, day in enumerate(DAYS, 2):
                slot = schedule[cls][day][period]
                if slot:
                    teacher_obj = next(
                        (t for t in teachers_list if t["id"] == slot["teacher_id"]),
                        None
                    )
                    t_name = teacher_obj["name"].split()[-1] if teacher_obj else ""
                    text   = f"{slot['subject']}\n({t_name})"
                    fill   = morning_fill if period in MORNING else afternoon_fill
                else:
                    text = "—"
                    fill = free_fill

                cell = ws.cell(row=p_i, column=d_i, value=text)
                style_cell(cell, fill=fill, font=cell_font)

    # ── Sheet: ตารางสอนรายครู ───────────────────────────────
    teacher_by_id = {t["id"]: t for t in teachers_list}
    for t_id, t_sched in teacher_schedules.items():
        t_name = teacher_by_id[t_id]["name"]
        ws = wb.create_sheet(title=f"ครู {teacher_by_id[t_id]['name'].split()[-1]}")
        ws.sheet_view.showGridLines = False

        ws.merge_cells("A1:I1")
        ws["A1"] = f"ตารางสอน — {t_name}  |  โรงเรียนวัดโป่งแรด"
        ws["A1"].fill      = header_fill
        ws["A1"].font      = Font(name="TH Sarabun New", bold=True, color="FFFFFF", size=15)
        ws["A1"].alignment = center
        ws.row_dimensions[1].height = 36

        ws.column_dimensions["A"].width = 6
        headers = ["คาบ"] + DAYS
        for col_i, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_i, value=h)
            cell.fill = subhead_fill
            style_cell(cell, font=header_font)
            ws.column_dimensions[get_column_letter(col_i)].width = 22

        for p_i, period in enumerate(PERIODS, 3):
            ws.cell(row=p_i, column=1, value=f"คาบ {period}")
            style_cell(ws.cell(row=p_i, column=1),
                       fill=PatternFill("solid", fgColor="E9ECEF"),
                       font=bold_font)
            ws.row_dimensions[p_i].height = 38

            period_count = 0
            for d_i, day in enumerate(DAYS, 2):
                slot = t_sched[day][period]
                if slot:
                    text  = f"{slot['subject']}\n({slot['class']})"
                    fill  = morning_fill if period in MORNING else afternoon_fill
                    period_count += 1
                else:
                    text = "— ว่าง —"
                    fill = free_fill

                cell = ws.cell(row=p_i, column=d_i, value=text)
                style_cell(cell, fill=fill, font=cell_font)

    # ── Sheet: สรุปภาพรวม ────────────────────────────────────
    ws = wb.create_sheet(title="สรุปภาพรวม", index=0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 5

    ws.merge_cells("B1:G1")
    ws["B1"] = "สรุปภาระงานสอน — โรงเรียนวัดโป่งแรด"
    ws["B1"].fill      = header_fill
    ws["B1"].font      = Font(name="TH Sarabun New", bold=True, color="FFFFFF", size=16)
    ws["B1"].alignment = center
    ws.row_dimensions[1].height = 40

    summary_headers = ["ชื่อครู", "วิชาหลัก", "คาบสอน/สัปดาห์", "คาบว่าง/สัปดาห์", "สถานะ"]
    for col_i, h in enumerate(summary_headers, 2):
        cell = ws.cell(row=2, column=col_i, value=h)
        cell.fill = subhead_fill
        style_cell(cell, font=header_font)
        ws.column_dimensions[get_column_letter(col_i)].width = 24

    for row_i, t in enumerate(teachers_list, 3):
        t_sched = teacher_schedules.get(t["id"], {})
        total_slots = 5 * 8  # 5 วัน × 8 คาบ
        used = sum(
            1 for day in DAYS for p in PERIODS
            if t_sched.get(day, {}).get(p)
        )
        free   = total_slots - used
        status = "✅ ปกติ" if free >= 5 else "⚠️ งานหนัก"

        row_data = [t["name"], t["specialty"], used, free, status]
        fill_row = PatternFill("solid", fgColor="F0FAF4") if row_i % 2 == 0 else None
        for col_i, val in enumerate(row_data, 2):
            cell = ws.cell(row=row_i, column=col_i, value=val)
            style_cell(cell, fill=fill_row, font=cell_font)
        ws.row_dimensions[row_i].height = 28

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── Session State Init ──────────────────────────────────────
if "teachers" not in st.session_state:
    st.session_state.teachers = deepcopy(DEFAULT_TEACHERS)
if "subjects" not in st.session_state:
    st.session_state.subjects = deepcopy(DEFAULT_SUBJECTS)
if "classes" not in st.session_state:
    st.session_state.classes = list(DEFAULT_CLASSES)
if "schedule" not in st.session_state:
    st.session_state.schedule = None
if "teacher_busy" not in st.session_state:
    st.session_state.teacher_busy = None
if "conflicts" not in st.session_state:
    st.session_state.conflicts = []
if "teacher_by_id" not in st.session_state:
    st.session_state.teacher_by_id = {}


# ─── Sidebar ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ ตั้งค่าระบบ")
    st.markdown("---")

    st.markdown("### 📅 Soft Constraints")
    soft_focus    = st.toggle("วิชา Focus อยู่ช่วงเช้า", value=True)
    soft_gap      = st.toggle("ครูมีคาบว่างอย่างน้อย 1 คาบ/วัน", value=True)
    soft_consec   = st.toggle("สอนวิชาเดิมต่อเนื่อง", value=True)

    st.markdown("---")
    st.markdown("### 🏫 ข้อมูลโรงเรียน")
    st.info(f"""
    👩‍🏫 ครู: **{len(st.session_state.teachers)} คน**  
    📚 วิชา: **{len(st.session_state.subjects)} วิชา**  
    🏠 ชั้นเรียน: **{len(st.session_state.classes)} ห้อง**  
    📅 วัน/สัปดาห์: **5 วัน**  
    🕐 คาบ/วัน: **8 คาบ**
    """)

    st.markdown("---")
    if st.button("🚀 สร้างตารางอัตโนมัติ", use_container_width=True):
        with st.spinner("กำลังประมวลผล..."):
            sched, busy, conflicts, t_by_id = generate_timetable(
                st.session_state.classes,
                st.session_state.subjects,
                st.session_state.teachers,
                soft_focus,
                soft_gap,
            )
            st.session_state.schedule    = sched
            st.session_state.teacher_busy = busy
            st.session_state.conflicts   = conflicts
            st.session_state.teacher_by_id = t_by_id

            teacher_schedules = {
                t["id"]: build_teacher_schedule(t["id"], sched, st.session_state.classes)
                for t in st.session_state.teachers
            }
            st.session_state.teacher_schedules = teacher_schedules
        st.success("✅ สร้างตารางเสร็จแล้ว!")

    if st.session_state.schedule:
        st.markdown("---")
        excel_buf = export_to_excel(
            st.session_state.schedule,
            st.session_state.get("teacher_schedules", {}),
            st.session_state.teachers,
            st.session_state.classes,
        )
        st.download_button(
            label="📥 Export Excel",
            data=excel_buf,
            file_name="ตารางสอน_วัดโป่งแรด.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ─── Main Content ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🏫 ระบบจัดตารางเรียนตารางสอนอัตโนมัติ</h1>
  <p>โรงเรียนวัดโป่งแรด · Timetable Management System</p>
</div>
""", unsafe_allow_html=True)

# Stats row
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f"""<div class="stat-card">
        <div class="stat-number">{len(st.session_state.teachers)}</div>
        <div class="stat-label">👩‍🏫 ครูทั้งหมด</div>
    </div>""", unsafe_allow_html=True)
with col2:
    st.markdown(f"""<div class="stat-card">
        <div class="stat-number">{len(st.session_state.subjects)}</div>
        <div class="stat-label">📚 วิชาทั้งหมด</div>
    </div>""", unsafe_allow_html=True)
with col3:
    st.markdown(f"""<div class="stat-card">
        <div class="stat-number">{len(st.session_state.classes)}</div>
        <div class="stat-label">🏠 ชั้นเรียน</div>
    </div>""", unsafe_allow_html=True)
with col4:
    n_conflicts = len(st.session_state.conflicts)
    color = "#dc3545" if n_conflicts > 0 else "#1a472a"
    st.markdown(f"""<div class="stat-card" style="border-left-color:{color}">
        <div class="stat-number" style="color:{color}">{n_conflicts}</div>
        <div class="stat-label">⚠️ Conflicts</div>
    </div>""", unsafe_allow_html=True)


# ─── Tabs ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 ตารางเรียนรายชั้น",
    "👩‍🏫 ตารางสอนรายครู",
    "✏️ จัดการข้อมูล",
    "🔍 ตรวจสอบ Conflicts",
])


# ── Tab 1: ตารางเรียนรายชั้น ─────────────────────────────────
with tab1:
    if not st.session_state.schedule:
        st.info("กด **🚀 สร้างตารางอัตโนมัติ** ที่แถบด้านซ้ายเพื่อเริ่มต้น")
    else:
        selected_class = st.selectbox(
            "เลือกชั้นเรียน",
            st.session_state.classes,
            key="class_view"
        )

        sched = st.session_state.schedule[selected_class]

        # Build display DataFrame
        rows = []
        for period in PERIODS:
            row = {"คาบ": f"คาบ {period}"}
            for day in DAYS:
                slot = sched[day][period]
                if slot:
                    t = st.session_state.teacher_by_id.get(slot["teacher_id"])
                    t_name = t["name"].split()[-1] if t else "?"
                    row[day] = f"{slot['subject']} ({t_name})"
                else:
                    row[day] = "—"
            rows.append(row)

        df = pd.DataFrame(rows).set_index("คาบ")

        # Color styling
        def color_cell(val):
            subj = val.split(" (")[0] if "(" in val else val
            color = SUBJECT_COLOR.get(subj, "#ffffff")
            return f"background-color: {color}; font-weight: 500;"

        st.markdown(f"#### ตารางเรียน ชั้น {selected_class}")
        st.dataframe(
            df.style.map(color_cell),
            use_container_width=True,
            height=360,
        )

        # Legend
        st.markdown("**สีวิชา:** ", unsafe_allow_html=True)
        legend_cols = st.columns(5)
        for i, subj in enumerate(st.session_state.subjects):
            with legend_cols[i % 5]:
                color = SUBJECT_COLOR.get(subj["name"], "#eee")
                st.markdown(
                    f'<span style="background:{color};padding:2px 8px;'
                    f'border-radius:4px;font-size:0.8rem">{subj["name"]}</span>',
                    unsafe_allow_html=True,
                )


# ── Tab 2: ตารางสอนรายครู ────────────────────────────────────
with tab2:
    if not st.session_state.schedule:
        st.info("กด **🚀 สร้างตารางอัตโนมัติ** ที่แถบด้านซ้ายเพื่อเริ่มต้น")
    else:
        teacher_names = [t["name"] for t in st.session_state.teachers]
        selected_t = st.selectbox("เลือกครู", teacher_names, key="teacher_view")
        t_obj = next(t for t in st.session_state.teachers if t["name"] == selected_t)
        t_sched = st.session_state.teacher_schedules.get(t_obj["id"], {})

        rows = []
        total_used = 0
        for period in PERIODS:
            row = {"คาบ": f"คาบ {period}"}
            for day in DAYS:
                slot = t_sched[day][period]
                if slot:
                    row[day] = f"{slot['subject']} ({slot['class']})"
                    total_used += 1
                else:
                    row[day] = "— ว่าง —"
            rows.append(row)

        df = pd.DataFrame(rows).set_index("คาบ")

        def color_teacher(val):
            if "ว่าง" in val:
                return "background-color: #f8f9fa; color: #adb5bd;"
            subj = val.split(" (")[0]
            color = SUBJECT_COLOR.get(subj, "#ffffff")
            return f"background-color: {color}; font-weight: 500;"

        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.markdown(f"#### ตารางสอน — {selected_t} ({t_obj['specialty']})")
        with col_b:
            free = 5 * 8 - total_used
            st.markdown(
                f"📊 สอน **{total_used}** คาบ | ว่าง **{free}** คาบ/สัปดาห์",
            )

        st.dataframe(
            df.style.map(color_teacher),
            use_container_width=True,
            height=360,
        )


# ── Tab 3: จัดการข้อมูล ─────────────────────────────────────
with tab3:
    sub1, sub2, sub3 = st.tabs(["👩‍🏫 ครู", "📚 วิชา", "🏠 ชั้นเรียน"])

    with sub1:
        st.markdown("#### รายชื่อครูทั้งหมด")
        teacher_df = pd.DataFrame(st.session_state.teachers)
        edited_teachers = st.data_editor(
            teacher_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
        )
        if st.button("💾 บันทึกข้อมูลครู"):
            st.session_state.teachers = edited_teachers.to_dict("records")
            st.success("บันทึกแล้ว!")

    with sub2:
        st.markdown("#### รายวิชาทั้งหมด")
        subject_df = pd.DataFrame(st.session_state.subjects)
        edited_subjects = st.data_editor(
            subject_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
        )
        if st.button("💾 บันทึกข้อมูลวิชา"):
            st.session_state.subjects = edited_subjects.to_dict("records")
            st.success("บันทึกแล้ว!")

    with sub3:
        st.markdown("#### ชั้นเรียนทั้งหมด")
        classes_input = st.text_area(
            "รายชื่อชั้น (บรรทัดละ 1 ชั้น)",
            "\n".join(st.session_state.classes),
            height=200,
        )
        if st.button("💾 บันทึกชั้นเรียน"):
            st.session_state.classes = [
                c.strip() for c in classes_input.split("\n") if c.strip()
            ]
            st.success("บันทึกแล้ว!")


# ── Tab 4: Conflicts ─────────────────────────────────────────
with tab4:
    st.markdown("#### 🔍 ผลการตรวจสอบ Constraints")

    if not st.session_state.schedule:
        st.info("ยังไม่ได้สร้างตาราง")
    else:
        conflicts = st.session_state.conflicts
        if not conflicts:
            st.success("✅ ไม่พบ Conflict ทั้งหมด! ตารางสอนสมบูรณ์")
        else:
            st.error(f"❌ พบ {len(conflicts)} Conflict")
            for c in conflicts:
                st.markdown(f"- ⚠️ {c}")

        # Teacher workload summary
        st.markdown("---")
        st.markdown("#### 📊 สรุปภาระงานสอนรายครู")

        summary_rows = []
        for t in st.session_state.teachers:
            t_sched = st.session_state.teacher_schedules.get(t["id"], {})
            used = sum(
                1 for day in DAYS for p in PERIODS
                if t_sched.get(day, {}).get(p)
            )
            free    = 5 * 8 - used
            status  = "✅ ปกติ" if free >= 5 else "⚠️ งานหนัก"
            summary_rows.append({
                "ชื่อครู":          t["name"],
                "วิชาหลัก":         t["specialty"],
                "คาบสอน/สัปดาห์": used,
                "คาบว่าง/สัปดาห์": free,
                "สถานะ":            status,
            })

        summary_df = pd.DataFrame(summary_rows)

        def highlight_status(val):
            if "⚠️" in str(val):
                return "background-color:#fff3cd; color:#856404; font-weight:600"
            if "✅" in str(val):
                return "background-color:#d1e7dd; color:#0f5132; font-weight:600"
            return ""

        st.dataframe(
            summary_df.style.map(highlight_status, subset=["สถานะ"]),
            use_container_width=True,
            hide_index=True,
        )
