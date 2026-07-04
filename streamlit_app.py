import streamlit as st
import sqlite3, time, calendar as cal_lib, glob, os, re, threading, json, io
import pandas as pd
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

DB = "salesdb.db"
TT_DIR = "인트라넷 시간표"
DAYS = "월화수목금토일"
KST = ZoneInfo("Asia/Seoul")

# ── 시간표 엑셀(HTML export) 파싱 ────────────────────────────────
CELL_RE = re.compile(
    r'^(?P<subject>.+?)\s*전체출석율\s*:\s*[\d.]+%.*?'
    r'정원\s*:\s*(?P<cap>\d+)\((?P<enrolled>\d+)\).*?'
    r'배정\s*:\s*(?P<assigned>\d+)(?:\(W:\d+,R:\d+\))?'
    r'(?P<rest>.+?)'
    r'개\s*:\s*(?P<start_date>\d{4}-\d{2}-\d{2})종\s*:\s*(?P<end_date>\d{4}-\d{2}-\d{2})\s*$'
)
DAY_TAIL_RE = re.compile(r'([월화수목금토일][월화수목금토일~,/]*)$')

def _end_plus30(t):
    h, m = map(int, t.split(":"))
    m += 30
    if m >= 60: m -= 60; h += 1
    return f"{h:02d}:{m:02d}"

def _expand_days(tok):
    tok = tok.strip()
    if "~" in tok:
        a, b = tok.split("~")
        if a in DAYS and b in DAYS:
            i, j = DAYS.index(a), DAYS.index(b)
            return list(DAYS[i:j+1])
        return []
    seen = []
    for ch in tok:
        if ch in DAYS and ch not in seen:
            seen.append(ch)
    return seen

def _parse_cell(text, room, time_label):
    m = CELL_RE.match(text.strip())
    if not m:
        return None
    rest = m.group("rest")
    dm = DAY_TAIL_RE.search(rest)
    if not dm:
        return None
    day_tok = dm.group(1)
    days = _expand_days(day_tok)
    if not days:
        return None
    teacher = re.sub(r"\d+$", "", rest[:dm.start()].strip())
    teacher = re.sub(r"^재직자\s*:\s*\d+", "", teacher).strip()
    subject = re.sub(r"/주말$", "", m.group("subject")).strip()
    return dict(room=room, subject=subject, teacher=teacher, days=days, day_label=day_tok,
                start_date=m.group("start_date"), end_date=m.group("end_date"),
                cap=int(m.group("cap")), enrolled=int(m.group("enrolled")),
                assigned=int(m.group("assigned")), start_time=time_label)

@st.cache_data
def _load_timetable(_cache_key):
    sessions = []
    room_order = []
    for path in sorted(glob.glob(os.path.join(TT_DIR, "*.xls"))):
        try:
            df = pd.read_html(path, header=[0, 1])[0]
        except Exception:
            continue
        df = df[df.iloc[:, 0] != "정원"].reset_index(drop=True)
        times = df.iloc[:, 0].tolist()
        for col in df.columns[1:]:
            room = col[0]
            if room not in room_order:
                room_order.append(room)
            vals = df[col].tolist()
            i = 0
            while i < len(vals):
                v = vals[i]
                if pd.isna(v):
                    i += 1; continue
                j = i
                while j + 1 < len(vals) and vals[j + 1] == v:
                    j += 1
                sess = _parse_cell(str(v), room, times[i])
                if sess:
                    sess["end_time"] = _end_plus30(times[j])
                    sessions.append(sess)
                i = j + 1
    return sessions, room_order

def get_timetable():
    files = glob.glob(os.path.join(TT_DIR, "*.xls"))
    key = tuple(sorted((os.path.basename(f), os.path.getmtime(f)) for f in files))
    return _load_timetable(key)[0]

def get_room_order():
    files = glob.glob(os.path.join(TT_DIR, "*.xls"))
    key = tuple(sorted((os.path.basename(f), os.path.getmtime(f)) for f in files))
    return _load_timetable(key)[1]

# ── 개인 시간표(양식 xlsx) 파싱 ───────────────────────────────────
def _parse_personal_cell(text):
    lines = [l.strip() for l in str(text).split("\n") if l.strip()]
    if not lines:
        return None
    start_date = end_date = None
    day_tok = None
    for l in lines:
        m = re.match(r"^개\s*:\s*(\d{4}-\d{2}-\d{2})$", l)
        if m: start_date = m.group(1)
        m = re.match(r"^종\s*:\s*(\d{4}-\d{2}-\d{2})$", l)
        if m: end_date = m.group(1)
        dm = re.match(r"^[월화수목금토일][월화수목금토일~,/]*$", l)
        if dm: day_tok = l
    if not (start_date and end_date and day_tok):
        return None
    days = _expand_days(day_tok)
    if not days:
        return None
    return dict(subject=lines[0], day_label=day_tok, days=days,
                start_date=start_date, end_date=end_date)

def parse_personal_timetable(file):
    import openpyxl
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active

    student = ""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                m = re.search(r"(\S+)님\s*개인\s*시간표", cell.value)
                if m:
                    student = m.group(1).strip()
                    break
        if student:
            break

    courses = []
    rows = list(ws.iter_rows(values_only=True))
    for row in rows:
        cells = list(row)
        if not cells or cells[0] == "비고":
            continue
        rest = [c for c in cells[1:] if c not in (None, "")]
        if not rest:
            continue
        # 월 라벨 행(예: "7월","6월"만 있는 행)은 건너뜀
        if all(isinstance(c, str) and re.match(r"^\d{1,2}월$", c) for c in rest):
            continue
        course_cells = [c for c in cells[1:] if isinstance(c, str) and "개:" in c]
        if not course_cells:
            continue
        time_lines = []
        if isinstance(cells[0], str):
            time_lines = [t.strip() for t in cells[0].split("\n") if t.strip()]
        for i, ctext in enumerate(course_cells):
            parsed = _parse_personal_cell(ctext)
            if not parsed:
                continue
            if len(time_lines) == len(course_cells):
                tl = time_lines[i]
            elif time_lines:
                tl = time_lines[0]
            else:
                tl = ""
            start_time = tl.split("~")[0].strip() if tl else ""
            parsed["start_time"] = start_time
            courses.append(parsed)
    return student, courses

# ── 시간표 이미지 인식 (Gemini vision) ────────────────────────────
@st.cache_resource
def get_gemini_client():
    api_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    if not api_key:
        return None
    from google import genai
    return genai.Client(api_key=api_key)

def parse_timetable_image(image_bytes, mime_type):
    client = get_gemini_client()
    if not client:
        raise RuntimeError("GEMINI_API_KEY가 설정되어 있지 않아요 (Streamlit Secrets 확인).")
    from google.genai import types
    prompt = (
        "이 이미지는 학원 개인 시간표야. 표 안의 각 강좌 블록에서 정보를 추출해서 "
        "JSON 배열로만 답해. 설명 문장은 절대 붙이지 마.\n"
        '형식: [{"subject":"과목명","day_label":"월~금","start_date":"YYYY-MM-DD",'
        '"end_date":"YYYY-MM-DD","start_time":"HH:MM"}]\n'
        "값을 모르면 빈 문자열로 둬."
    )
    from google.genai import errors as genai_errors
    parts = [types.Part.from_bytes(data=image_bytes, mime_type=mime_type), prompt]
    for attempt in range(3):
        try:
            resp = client.models.generate_content(model="gemini-2.5-flash-lite", contents=parts)
            break
        except genai_errors.ServerError:
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))  # ponytail: fixed backoff, exponential if 503s get frequent
    m = re.search(r"\[.*\]", resp.text.strip(), re.S)
    return json.loads(m.group(0)) if m else []

def sessions_on_date(d):
    wd = DAYS[d.weekday()]  # Mon=0..Sun=6 -> 월..일
    ds = d.isoformat()
    out = [s for s in get_timetable()
           if wd in s["days"] and s["start_date"] <= ds <= s["end_date"]]
    return sorted(out, key=lambda s: (s["start_time"], s["room"]))

# ── DB ───────────────────────────────────────────────────────
@st.cache_resource
def get_db():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.executescript("""
    CREATE TABLE IF NOT EXISTS tasks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        category TEXT DEFAULT '주중업무',
        task_date TEXT DEFAULT '',
        assignee TEXT DEFAULT '',
        is_done INTEGER DEFAULT 0,
        created_at INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS consultations(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, sched_date TEXT NOT NULL,
        sched_time TEXT DEFAULT '', expected_revenue INTEGER DEFAULT 0,
        ctype TEXT DEFAULT '단과',
        created_at INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS daily_log(
        log_date TEXT PRIMARY KEY,
        team_name TEXT DEFAULT '2-3팀',
        rep1_name TEXT DEFAULT '', rep1_pct INTEGER DEFAULT 60,
        rep2_name TEXT DEFAULT '',
        rep1_call TEXT DEFAULT '', rep2_call TEXT DEFAULT '',
        done_count INTEGER DEFAULT 0,
        registered INTEGER DEFAULT 0, cod INTEGER DEFAULT 0, unregistered INTEGER DEFAULT 0,
        actual_revenue INTEGER DEFAULT 0, refund INTEGER DEFAULT 0,
        interview_count INTEGER DEFAULT 0,
        ddaz_num INTEGER DEFAULT 0, ddaz_den INTEGER DEFAULT 32,
        tmr_target INTEGER DEFAULT 0,
        month_target INTEGER DEFAULT 0, month_achieved INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS app_state(
        key TEXT PRIMARY KEY,
        value TEXT);
    """)
    c.commit()
    try:
        c.execute("ALTER TABLE consultations ADD COLUMN ctype TEXT DEFAULT '단과'")
        c.commit()
    except sqlite3.OperationalError:
        pass  # 이미 있음
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN assignee TEXT DEFAULT ''")
        c.commit()
    except sqlite3.OperationalError:
        pass  # 이미 있음
    return c

# ponytail: global lock, connection pooling if concurrent traffic becomes a bottleneck
_db_lock = threading.Lock()

def q(sql, a=()):
    with _db_lock:
        return get_db().execute(sql, a).fetchall()

def run(sql, a=()):
    with _db_lock:
        get_db().execute(sql, a)
        get_db().commit()

def get_state(key, default=None):
    rows = q("SELECT value FROM app_state WHERE key=?", (key,))
    return rows[0][0] if rows else default

def set_state(key, value):
    run("INSERT INTO app_state(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value))

def get_tasks():
    rows = q("SELECT id,title,category,task_date,assignee,is_done FROM tasks ORDER BY is_done,created_at DESC")
    return [dict(zip("id title category task_date assignee is_done".split(), r)) for r in rows]

def get_consults():
    rows = q("SELECT id,name,sched_date,sched_time,expected_revenue,ctype FROM consultations ORDER BY sched_date,sched_time")
    return [dict(zip("id name sched_date sched_time expected_revenue ctype".split(), r)) for r in rows]

def add_task(title, cat, d, assignee):
    run("INSERT INTO tasks(title,category,task_date,assignee,is_done,created_at) VALUES(?,?,?,?,0,?)",
        (title, cat, d, assignee, int(time.time()*1000)))

def toggle_task(tid, v): run("UPDATE tasks SET is_done=? WHERE id=?", (int(v), tid))
def del_task(tid):       run("DELETE FROM tasks WHERE id=?", (tid,))

def add_consult(name, d, t, rev, ctype):
    run("INSERT INTO consultations(name,sched_date,sched_time,expected_revenue,ctype,created_at) VALUES(?,?,?,?,?,?)",
        (name, d, t, rev, ctype, int(time.time()*1000)))
def del_consult(cid): run("DELETE FROM consultations WHERE id=?", (cid,))

# ── 일지(daily_log): 자동계산 외 수동 항목 저장 ────────────────────
LOG_COLS = ("log_date team_name rep1_name rep1_pct rep2_name rep1_call rep2_call "
            "done_count registered cod unregistered actual_revenue refund "
            "interview_count ddaz_num ddaz_den tmr_target month_target month_achieved").split()

def get_log(d):
    rows = q(f"SELECT {','.join(LOG_COLS)} FROM daily_log WHERE log_date=?", (d,))
    if rows:
        return dict(zip(LOG_COLS, rows[0]))
    prev = q(f"SELECT {','.join(LOG_COLS)} FROM daily_log ORDER BY log_date DESC LIMIT 1")
    base = dict(zip(LOG_COLS, prev[0])) if prev else {}
    new_row = dict(log_date=d, team_name=base.get("team_name", "2-3팀"),
                    rep1_name=base.get("rep1_name", ""), rep1_pct=base.get("rep1_pct", 60),
                    rep2_name=base.get("rep2_name", ""), rep1_call="", rep2_call="",
                    done_count=0, registered=0, cod=0, unregistered=0, actual_revenue=0, refund=0,
                    interview_count=0, ddaz_num=0, ddaz_den=base.get("ddaz_den", 32),
                    tmr_target=0, month_target=base.get("month_target", 0), month_achieved=0)
    run("""INSERT INTO daily_log(log_date, team_name, rep1_name, rep1_pct, rep2_name, ddaz_den, month_target)
           VALUES(?,?,?,?,?,?,?)""",
        (d, new_row["team_name"], new_row["rep1_name"], new_row["rep1_pct"],
         new_row["rep2_name"], new_row["ddaz_den"], new_row["month_target"]))
    return new_row

# ── 배정 문구 생성기 ──────────────────────────────────────────
def fdate(iso):
    if not iso: return ""
    p = iso.split("-")
    return f"{int(p[1])}.{int(p[2])}"

def gen_text(type_, **k):
    nd, nt = fdate(k.get("nd","")), k.get("nt","")
    ns, nts = k.get("ns",""), k.get("nts","")
    ns_full = f"{ns}/{nts}" if nts else ns
    nfee, ofee = k.get("nfee",0), k.get("ofee",0)
    od, ot, os_ = fdate(k.get("od","")), k.get("ot",""), k.get("os_","")

    if type_ == "신규":    return f"미배정 -> {nd} {nt} {ns_full} 배정"
    if type_ == "과목변경":
        if ofee or nfee:
            diff = abs(nfee - ofee)
            sfx  = f"\n[차액 {diff:,}원 무시]" if diff else ""
            return f"미배정 {os_} (수강료:{ofee:,}원) -> {nd} {nt} {ns_full} (수강료:{nfee:,}원) 과목변경 배정{sfx}"
        return f"미배정 {os_} -> {nd} {nt} {ns_full} 과목변경 배정"
    if type_ == "취소":    return f"{od} {ot} {os_} 배정 -> 미배정"
    if type_ == "날짜변경": return f"{od} {ot} {os_} 배정 -> {nd} {nt} {os_} 배정"

# ── 페이지 설정 ───────────────────────────────────────────────
st.set_page_config(page_title="업무 대시보드", page_icon="📊",
                   layout="wide", initial_sidebar_state="collapsed")

ss = st.session_state
if "assign_out" not in ss: ss.assign_out = ""

now_kst   = datetime.now(KST)
today     = now_kst.date()
today_str = today.isoformat()
yr, mo    = today.year, today.month

# 한국시간 06:00 기준으로 하루 업무일 판단 → 새 업무일이면 체크 초기화
workday = (today - timedelta(days=1)).isoformat() if now_kst.hour < 6 else today_str
if get_state("last_task_reset") != workday:
    run("UPDATE tasks SET is_done=0")
    set_state("last_task_reset", workday)

# ── 데이터 로드 ────────────────────────────────────────────────
tasks    = get_tasks()
consults = get_consults()
today_c  = [c for c in consults if c["sched_date"] == today_str]

tomorrow_str = (today + timedelta(days=1)).isoformat()
dayafter_str = (today + timedelta(days=2)).isoformat()

def day_stats(d):
    items = [c for c in consults if c["sched_date"] == d]
    cnt = len(items)
    rev = sum(c["expected_revenue"] for c in items)
    reg = sum(1 for c in items if c["ctype"] == "정규")
    return cnt, rev, reg, cnt - reg

today_cnt, today_rev, today_reg, today_dan = day_stats(today_str)
tmr_cnt, tmr_rev, _, _ = day_stats(tomorrow_str)
daf_cnt, daf_rev, _, _ = day_stats(dayafter_str)
log = get_log(today_str)

# 달력용 이벤트 맵
cmap = {}
for c in consults:
    if c["sched_date"][:7] == f"{yr:04d}-{mo:02d}":
        d = int(c["sched_date"][8:])
        cmap.setdefault(d, []).append(c)

# ── CSS (사용자 HTML 기반) ─────────────────────────────────────
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');

#MainMenu,footer,[data-testid="stHeader"],[data-testid="stToolbar"]{display:none!important}
.block-container{padding:.8rem 1.8rem 2rem!important;max-width:100%!important}
[data-testid="stAppViewContainer"],.main{background:#f8f9fa!important}

:root{
  --bg:#f8f9fa;--card:#fff;--tx:#333;--sub:#666;--bd:#eaeaea;
  --p1:#fff0f0;--p1t:#ff5b5b;
  --p2:#fff5e6;--p2t:#ff9f43;
  --p3:#f0f7ff;--p3t:#2f80ed;
}
*{font-family:'Noto Sans KR',sans-serif}

/* Streamlit 버튼 */
[data-testid="stButton"]>button{
  padding:3px 10px!important;font-size:11px!important;border-radius:5px!important;
  font-weight:500!important;min-height:0!important;line-height:1.5!important;
  border:1px solid var(--bd)!important;background:transparent!important;color:var(--sub)!important}
[data-testid="stButton"]>button:hover{background:#f0f0f0!important}
[data-testid="stCheckbox"] label{font-size:11px!important;color:var(--sub)!important}
[data-testid="stCheckbox"]>label>div:first-child{width:13px!important;height:13px!important;border-radius:3px!important}
[data-testid="stForm"]{border:none!important;padding:0!important}
.stTextInput input,.stSelectbox>div>div,.stDateInput input,.stNumberInput input{
  border-radius:7px!important;border:1.5px solid var(--bd)!important;font-size:12px!important;background:#fff!important}
[data-testid="stExpander"]{border:1.5px solid var(--bd)!important;border-radius:10px!important;background:#fff!important}
div[data-testid="stHorizontalBlock"]{gap:.8rem!important}
hr{margin:.3rem 0!important;border-color:#f0f0f0!important}

/* 공통 카드 */
.db-card{background:var(--card);border-radius:16px;padding:22px;
         box-shadow:0 4px 12px rgba(0,0,0,0.03);border:1px solid var(--bd)}
.db-card-title{font-size:16px;font-weight:700;color:var(--tx);
               display:flex;align-items:center;gap:7px;margin-bottom:16px}

/* 달력 */
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);
          border-top:1px solid var(--bd);border-left:1px solid var(--bd);
          border-radius:0 0 8px 8px;overflow:hidden}
.cal-head{background:#fdfdfd;padding:8px 4px;text-align:center;font-weight:700;
          font-size:12px;border-right:1px solid var(--bd);border-bottom:1px solid var(--bd);color:var(--sub)}
.cal-day{min-height:82px;padding:6px;border-right:1px solid var(--bd);
         border-bottom:1px solid var(--bd);background:#fff;vertical-align:top}
.cal-day.td{background:#f5f3ff}
.cal-num{font-size:12px;font-weight:500;margin-bottom:3px;display:inline-block;
         width:21px;height:21px;line-height:21px;text-align:center;border-radius:50%}
.cal-num.td{background:#4f46e5;color:#fff!important}
.ci{font-size:9px;padding:2px 5px;border-radius:4px;margin-bottom:2px;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:600;display:block}
.ci1{background:var(--p1);color:var(--p1t)}.ci2{background:var(--p2);color:var(--p2t)}.ci3{background:var(--p3);color:var(--p3t)}

/* 칸반 */
.k-head{font-size:12px;font-weight:700;padding:5px 11px;border-radius:6px;display:inline-block;margin-bottom:10px}
.kp1{background:var(--p1);color:var(--p1t)}.kp2{background:var(--p2);color:var(--p2t)}.kp3{background:var(--p3);color:var(--p3t)}
.t-card{background:#fff;border:1px solid var(--bd);border-radius:10px;
        padding:12px;margin-bottom:4px;box-shadow:0 2px 6px rgba(0,0,0,0.02);
        transition:transform .15s}
.t-card:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,0.05)}
.t-date{font-size:11px;color:var(--sub);margin-bottom:4px}
.t-title{font-size:13px;font-weight:500;line-height:1.45;color:var(--tx)}
.t-title.done{text-decoration:line-through;color:#bbb}
</style>""", unsafe_allow_html=True)

# ── 헤더 (실시간 시계, 한국시간 기준) ────────────────────────────
st.iframe("""
<html><body style="margin:0;background:#f8f9fa">
<div style="font-family:'Noto Sans KR',sans-serif;display:flex;justify-content:space-between;
            align-items:center;padding:4px 2px 14px 2px;border-bottom:2px solid #eaeaea">
  <h1 style="font-size:22px;font-weight:700;color:#333;margin:0">
    <span style="color:#4f46e5">■</span> 업무 및 일정 관리 대시보드
  </h1>
  <div id="live-clock" style="font-size:13px;color:#666"></div>
</div>
<script>
function updateClock() {
  var now = new Date();
  var dOpts = {timeZone:'Asia/Seoul', year:'numeric', month:'long', day:'numeric', weekday:'short'};
  var dateStr = now.toLocaleDateString('ko-KR', dOpts);
  var timeStr = now.toLocaleTimeString('ko-KR', {timeZone:'Asia/Seoul', hour12:false});
  document.getElementById('live-clock').innerHTML =
    dateStr + '&nbsp;<span style="font-size:22px;font-weight:200;font-family:monospace;color:#333">' + timeStr + '</span>';
}
updateClock();
setInterval(updateClock, 1000);
</script>
</body></html>
""", height=70)

# ── 메인 2컬럼 ────────────────────────────────────────────────
left, right = st.columns([1.05, 1.35], gap="medium")

# ── LEFT: 달력 ───────────────────────────────────────────────
with left:
    ci_cls = ["ci1","ci2","ci3"]
    day_clr = [None,None,None,None,None,"#2f80ed","#ff5b5b"]

    def pill_html(c, cls):
        rev = f"{c['expected_revenue']//10000}만" if c["expected_revenue"] >= 10000 \
              else (f"{c['expected_revenue']}원" if c["expected_revenue"] else "")
        sub = " ".join(x for x in [c["sched_time"], rev] if x)
        tip = f"{c['name']} {c['sched_time']} {c['expected_revenue']:,}원 ({c['ctype']})"
        return f"""<span class="ci {cls}" title="{tip}">
<b>{c['name']}</b>{' · '+sub if sub else ''}</span>"""

    rows_html = ""
    for week in cal_lib.monthcalendar(yr, mo):
        rows_html += "<tr style='display:contents'>"
        for wi, day in enumerate(week):
            if day == 0:
                rows_html += '<div class="cal-day" style="background:#fafafa"></div>'; continue
            is_td   = day == today.day
            td_cls  = " td" if is_td else ""
            nc      = "td" if is_td else ""
            dc      = day_clr[wi]
            nstyle  = f"color:{dc}" if dc and not is_td else ""
            events  = cmap.get(day, [])
            evhtml  = "".join(
                pill_html(e, ci_cls[i%3])
                for i,e in enumerate(events[:3])
            )
            rows_html += f"""
<div class="cal-day{td_cls}">
  <span class="cal-num {nc}" style="{nstyle}">{day}</span>
  {evhtml}
</div>"""

    st.markdown(f"""
<div class="db-card">
  <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:14px'>
    <div class="db-card-title" style="margin:0">🗓️ 캘린더 일정</div>
    <div style='font-size:13px;font-weight:600;color:#666'>{yr}년 {mo}월</div>
  </div>
  <div class="cal-grid">
    <div class="cal-head">월</div><div class="cal-head">화</div>
    <div class="cal-head">수</div><div class="cal-head">목</div>
    <div class="cal-head">금</div>
    <div class="cal-head" style="color:#2f80ed">토</div>
    <div class="cal-head" style="color:#ff5b5b">일</div>
    {rows_html}
  </div>
</div>""", unsafe_allow_html=True)

    # 상담 추가 폼
    with st.expander("＋ 상담 일정 추가"):
        with st.form("cf", clear_on_submit=True):
            r1 = st.columns(2)
            cname = r1[0].text_input("이름 *")
            cdate = r1[1].date_input("날짜", value=today)
            r2 = st.columns([1, 1, 1])
            ctime  = r2[0].text_input("시간", placeholder="14:00")
            crev   = r2[1].number_input("예정매출(원)", min_value=0, step=10000)
            ctype  = r2[2].selectbox("구분", ["단과", "정규"])
            if st.form_submit_button("저장", use_container_width=True) and cname.strip():
                add_consult(cname.strip(), cdate.isoformat(), ctime, int(crev), ctype); st.rerun()

# ── RIGHT: 상담 일정(날짜별) + 칸반 TO DO LIST ──────────────────
with right:
    upcoming = sorted((c for c in consults if c["sched_date"] >= today_str),
                       key=lambda c: (c["sched_date"], c["sched_time"]))
    by_date = {}
    for c in upcoming:
        by_date.setdefault(c["sched_date"], []).append(c)

    st.markdown('<div class="db-card" style="margin-bottom:14px"><div class="db-card-title">🔵 상담 일정</div>',
                unsafe_allow_html=True)
    if not by_date:
        st.markdown('<div style="color:#ccc;font-size:11px;text-align:center;padding:12px 0">예정된 상담 없음</div>',
                    unsafe_allow_html=True)
    for d in sorted(by_date):
        dd  = datetime.strptime(d, "%Y-%m-%d")
        wtag = ["월","화","수","목","금","토","일"][dd.weekday()]
        label = f"{dd.month}월 {dd.day}일 ({wtag}){' · 오늘' if d == today_str else ''}"
        clr = "#4f46e5" if d == today_str else "#999"
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{clr};margin:6px 0 4px'>{label}</div>",
                    unsafe_allow_html=True)
        for c in by_date[d]:
            cc1, cc2 = st.columns([1, 0.13])
            cc1.markdown(f"""
<div style='background:#eef2ff;border-left:3px solid #4f46e5;border-radius:6px;
            padding:7px 12px;margin:3px 0;display:flex;justify-content:space-between;align-items:center'>
  <span style='font-size:12px;font-weight:700;color:#3730a3'>{c['name']} <span style='font-weight:500;color:#8b8bd8'>({c['ctype']})</span></span>
  <span style='font-size:11px;color:#6366f1'>{c['sched_time']}</span>
  <span style='font-size:11px;font-weight:700;color:#4f46e5'>{c['expected_revenue']:,}원</span>
</div>""", unsafe_allow_html=True)
            if cc2.button("✕", key=f"dc{c['id']}"):
                del_consult(c["id"]); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="db-card"><div class="db-card-title">☑ TO DO LIST</div>', unsafe_allow_html=True)
    st.caption("매일 한국시간 06:00에 완료 체크가 초기화돼요.")

    CATS = [
        ("우선순위1","우선순위 1","kp1"),
        ("우선순위2","우선순위 2","kp2"),
        ("주중업무", "주중업무",  "kp3"),
    ]
    assignees = sorted({t["assignee"] for t in tasks if t["assignee"]})
    groups = assignees + ["미지정"]

    for person in groups:
        person_tasks = [t for t in tasks if (t["assignee"] or "미지정") == person]
        left_n = sum(1 for t in person_tasks if not t["is_done"])
        with st.expander(f"👤 {person} ({left_n}건 남음)", expanded=True):
            k1, k2, k3 = st.columns(3, gap="small")

            for col_w, (cat_key, cat_label, kp_cls) in zip([k1, k2, k3], CATS):
                with col_w:
                    st.markdown(f'<div class="k-head {kp_cls}">{cat_label}</div>', unsafe_allow_html=True)

                    cat_tasks = [t for t in person_tasks if t["category"] == cat_key]
                    if not cat_tasks:
                        st.markdown('<div style="color:#ccc;font-size:11px;text-align:center;padding:12px 0">할 일 없음</div>',
                                    unsafe_allow_html=True)

                    for t in cat_tasks:
                        is_done   = bool(t["is_done"])
                        d_display = t["task_date"][5:].replace("-",".") if t["task_date"] else ""
                        tc        = "done" if is_done else ""

                        st.markdown(f"""
<div class="t-card">
  {'<div class="t-date">'+d_display+'</div>' if d_display else ''}
  <div class="t-title {tc}">{t['title']}</div>
</div>""", unsafe_allow_html=True)

                        c1, c2 = st.columns([2.5, 0.8])
                        new_done = c1.checkbox("완료", value=is_done, key=f"t{t['id']}")
                        if new_done != is_done:
                            toggle_task(t["id"], new_done); st.rerun()
                        if c2.button("✕", key=f"d{t['id']}"):
                            del_task(t["id"]); st.rerun()

                    # 할 일 추가 폼 (이 섹션 담당자로 자동 배정, 필요하면 수정 가능)
                    with st.form(f"af_{person}_{cat_key}", clear_on_submit=True):
                        new_title = st.text_input("", placeholder="＋ 새 페이지",
                                                  label_visibility="collapsed", key=f"ti_{person}_{cat_key}")
                        r = st.columns(2)
                        nd_val = r[0].date_input("날짜", value=today,
                                               label_visibility="collapsed", key=f"nd_{person}_{cat_key}")
                        new_assignee = r[1].text_input("담당자", value=("" if person == "미지정" else person),
                                               placeholder="담당자", label_visibility="collapsed",
                                               key=f"as_{person}_{cat_key}")
                        if st.form_submit_button("추가", use_container_width=True) and new_title.strip():
                            add_task(new_title.strip(), cat_key, nd_val.isoformat(), new_assignee.strip()); st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 시간표 배정 자동화 ─────────────────────────────────────────
st.markdown("""
<div class="db-card">
  <div class="db-card-title">📋 시간표 배정 자동화</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    정보 입력 → 배정 문구 자동 생성 후 복사
  </div>
</div>""", unsafe_allow_html=True)

candidates = []  # 체크박스로 선택된 강좌 → 아래 "선택한 강좌 배정 문구 생성"에서 일괄 처리
if "cb_gen_seq" not in st.session_state:
    st.session_state["cb_gen_seq"] = 0
cb_seq = st.session_state["cb_gen_seq"]  # 생성 후 체크박스 초기화(재마운트)용

with st.expander("🔍 날짜 클릭 → 그날 시간표에서 강좌 선택 (신규 배정용)", expanded=True):
    pick_date = st.date_input("날짜", value=today, key="tt_pick_date")
    day_sessions = sessions_on_date(pick_date)
    if day_sessions:
        by_room = {}
        for s in day_sessions:
            by_room.setdefault(s["room"], []).append(s)
        rooms = [r for r in get_room_order() if r in by_room]

        ROOMS_PER_ROW = 6
        for start in range(0, len(rooms), ROOMS_PER_ROW):
            row_rooms = rooms[start:start + ROOMS_PER_ROW]
            cols = st.columns(len(row_rooms))
            for col, room in zip(cols, row_rooms):
                with col:
                    st.markdown(f"""<div style='font-size:11px;font-weight:700;text-align:center;
                                background:#eef2ff;color:#4f46e5;border-radius:6px;padding:4px;margin-bottom:6px'>
                                {room}</div>""", unsafe_allow_html=True)
                    for si, s in enumerate(sorted(by_room[room], key=lambda x: x["start_time"])):
                        label = f"{s['start_time']} {s['subject']} ({s['teacher']})"
                        key = f"ttpick_{room}_{si}_{cb_seq}"
                        st.checkbox(label, key=key)
                        candidates.append(dict(key=key, nd=pick_date.isoformat(),
                                                nt=s["start_time"], ns=s["subject"], nts=s["day_label"]))
    else:
        st.caption("이 날짜에 진행 중인 강좌가 없어요.")

with st.expander("📎 개인 시간표 업로드해서 강좌 선택 (신규 배정용)"):
    upl = st.file_uploader("개인 시간표(xlsx)", type=["xlsx"], key="personal_tt_upl")
    if upl:
        try:
            student, courses = parse_personal_timetable(upl)
        except Exception as e:
            student, courses = "", []
            st.error(f"파일을 읽지 못했어요: {e}")
        if student:
            st.caption(f"{student}님의 시간표에서 {len(courses)}개 강좌를 찾았어요.")
        if courses:
            for i, c in enumerate(courses):
                label = f"{c['start_time']} {c['subject']} [{c['day_label']}] 개강 {c['start_date']}"
                key = f"pt_pick_{i}_{cb_seq}"
                st.checkbox(label, key=key)
                candidates.append(dict(key=key, nd=c["start_date"], nt=c["start_time"],
                                        ns=c["subject"], nts=c["day_label"]))
        elif upl and not student:
            st.caption("강좌 정보를 찾지 못했어요. 파일 형식을 확인해주세요.")

with st.expander("🖼️ 시간표 이미지 업로드 (AI 인식, 실험적)"):
    st.caption("AI가 이미지를 읽어서 추출하는 거라 100% 정확하진 않아요. 채워진 값을 확인하고 쓰세요.")

    from streamlit_paste_button import paste_image_button
    if "tt_img_seq" not in st.session_state:
        st.session_state["tt_img_seq"] = 0
    seq = st.session_state["tt_img_seq"]

    paste_result = paste_image_button("📋 클립보드에서 붙여넣기", key=f"tt_img_paste_{seq}")
    img = st.file_uploader("또는 파일로 업로드", type=["png", "jpg", "jpeg"], key=f"tt_img_upl_{seq}")

    if paste_result.image_data is not None:
        buf = io.BytesIO()
        paste_result.image_data.save(buf, format="PNG")
        st.session_state["tt_img_bytes"] = buf.getvalue()
        st.session_state["tt_img_mime"] = "image/png"
        st.session_state["tt_img_seq"] += 1  # 다음 붙여넣기가 되도록 컴포넌트 재마운트
    elif img:
        st.session_state["tt_img_bytes"] = img.getvalue()
        st.session_state["tt_img_mime"] = img.type

    image_bytes = st.session_state.get("tt_img_bytes")
    mime_type = st.session_state.get("tt_img_mime", "image/png")

    if image_bytes:
        st.image(image_bytes, caption="붙여넣은(또는 업로드한) 이미지", width=300)

    if image_bytes and st.button("이미지에서 강좌 인식하기", key="tt_img_go"):
        try:
            st.session_state["tt_img_results"] = parse_timetable_image(image_bytes, mime_type)
        except Exception as e:
            st.session_state["tt_img_results"] = []
            st.error(f"인식 실패: {e}")

    img_results = st.session_state.get("tt_img_results")
    if img_results:
        for i, c in enumerate(img_results):
            label = f"{c.get('start_time','')} {c.get('subject','')} [{c.get('day_label','')}] 개강 {c.get('start_date','')}"
            key = f"img_pick_{i}_{cb_seq}"
            st.checkbox(label, key=key)
            candidates.append(dict(key=key, nd=c.get("start_date", ""), nt=c.get("start_time", ""),
                                    ns=c.get("subject", ""), nts=c.get("day_label", "")))
    elif img_results is not None:
        st.caption("인식된 강좌가 없어요.")

if candidates:
    if st.button("✅ 체크한 강좌 배정 문구 생성 (여러 개 누적 가능)", use_container_width=True, key="gen_checked"):
        picked = [c for c in candidates if st.session_state.get(c["key"])]
        if picked:
            lines = [gen_text("신규", nd=c["nd"] or today.isoformat(), nt=c["nt"], ns=c["ns"], nts=c["nts"])
                      for c in picked]
            new_text = "\n".join(lines)
            ss.assign_out = (ss.assign_out + "\n" + new_text) if ss.assign_out else new_text
            st.session_state["cb_gen_seq"] += 1  # 체크박스 전부 재마운트해 초기화 (중복 누적 방지)
            st.rerun()
        else:
            st.caption("체크된 강좌가 없어요.")

gtype = st.selectbox("배정 유형", ["신규 배정","과목변경 배정","배정 취소","날짜변경 배정"],
                     label_visibility="collapsed")
result = ""

if gtype == "신규 배정":
    with st.form("g1"):
        c1,c2,c3,c4 = st.columns(4)
        nd  = c1.date_input("날짜*", value=today, key="g1_nd")
        nt  = c2.text_input("시간*", placeholder="12:00", key="g1_nt")
        ns  = c3.text_input("과목명*", placeholder="스케치업2", key="g1_ns")
        nts = c4.text_input("시간대", placeholder="주말", key="g1_nts")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("신규", nd=nd.isoformat(), nt=nt, ns=ns, nts=nts)

elif gtype == "과목변경 배정":
    with st.form("g2"):
        st.markdown("<div style='font-size:11px;color:#999;margin-bottom:4px'>▸ 이전</div>", unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        os_  = c1.text_input("이전 과목명", placeholder="캐드2")
        ofee = c2.number_input("이전 수강료(0=없음)", min_value=0, step=10000)
        st.markdown("<div style='font-size:11px;color:#999;margin:8px 0 4px'>▸ 변경 후</div>", unsafe_allow_html=True)
        c3,c4,c5,c6 = st.columns(4)
        nd  = c3.date_input("날짜*", value=today)
        nt  = c4.text_input("시간*", placeholder="19:00")
        ns  = c5.text_input("과목명*", placeholder="실내건축이론1")
        nts = c6.text_input("시간대", placeholder="주말")
        nfee = st.number_input("새 수강료(0=없음)", min_value=0, step=10000)
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("과목변경", os_=os_, ofee=int(ofee),
                              nd=nd.isoformat(), nt=nt, ns=ns, nts=nts, nfee=int(nfee))

elif gtype == "배정 취소":
    with st.form("g3"):
        c1,c2,c3 = st.columns(3)
        od  = c1.date_input("날짜*", value=today)
        ot  = c2.text_input("시간*", placeholder="19:00")
        os_ = c3.text_input("과목명*", placeholder="실내건축이론1")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("취소", od=od.isoformat(), ot=ot, os_=os_)

elif gtype == "날짜변경 배정":
    with st.form("g4"):
        st.markdown("<div style='font-size:11px;color:#999;margin-bottom:4px'>▸ 이전</div>", unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        od  = c1.date_input("날짜", value=today, key="od")
        ot  = c2.text_input("시간", placeholder="19:00", key="ot")
        os_ = c3.text_input("과목명", placeholder="캐드2")
        st.markdown("<div style='font-size:11px;color:#999;margin:8px 0 4px'>▸ 새 일정</div>", unsafe_allow_html=True)
        c4,c5 = st.columns(2)
        nd = c4.date_input("날짜", value=today, key="nd")
        nt = c5.text_input("시간", placeholder="19:00", key="nt")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("날짜변경", od=od.isoformat(), ot=ot, os_=os_,
                              nd=nd.isoformat(), nt=nt)

if result: ss.assign_out = result
if ss.assign_out:
    lines_n = ss.assign_out.count("\n") + 1
    st.text_area("📋 생성된 배정 문구 (복사하세요)", value=ss.assign_out, height=max(80, min(400, 30 * lines_n)))
    if st.button("🗑 초기화"):
        ss.assign_out = ""; st.rerun()

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 영업일지 자동 생성 (달력 상담/매출 자동계산, 나머지는 직접 수정) ─
st.markdown("""
<div class="db-card">
  <div class="db-card-title">📈 영업일지 자동 생성</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    상담건수·매출은 캘린더 기준으로 자동 채워집니다. 나머지 값은 아래 칸에서 바로 수정해서 쓰세요.
  </div>
</div>""", unsafe_allow_html=True)

if "rt_seq" not in st.session_state:
    st.session_state["rt_seq"] = 0
if st.button("🔄 캘린더 최신값으로 새로고침 (직접 수정한 내용은 초기화돼요)"):
    st.session_state["rt_seq"] += 1
    st.rerun()
rt_seq = st.session_state["rt_seq"]

pct = round(log['month_achieved'] / log['month_target'] * 100) if log['month_target'] else 0

morning_default = f"""{log['team_name']} 영업일지({today.month:02d}.{today.day:02d})

- 금일 입금예정: {today_rev // 10000}만원

{today_rev // 10000}만원 / {log['rep1_name']} {log['rep1_pct']}%

- 금일 상담건수 : {today_cnt}건(정규{today_reg}건/단과{today_dan}건)
- 면접예정: {log['interview_count']}건"""

pm3_default = f"""[{log['team_name']} 15:00 보고]
{log['done_count']} / {today_cnt}
익일상담 {tmr_cnt} / 익일예정 {tmr_rev // 10000}
모레상담 {daf_cnt} / 모레예정 {daf_rev // 10000}
익일면접 {log['interview_count']}건
따즈아 {log['ddaz_num']} / {log['ddaz_den']}"""

close_default = f"""컴퓨터 {log['team_name']} 영업마감보고

 상담 : {today_cnt}
 등록 : {log['registered']}
 COD : {log['cod']}
 미등록 : {log['unregistered']}

금일매출결과 :{log['actual_revenue']:,}원
환불 :{log['refund']:,}원

통화시간
{log['rep1_name']} {log['rep1_call']} (상담 {today_cnt}건)
{log['rep2_name']} {log['rep2_call']}

익일예정상담 : {tmr_cnt}건
익일예정매출 : {tmr_rev:,}원
익일목표매출 : {log['tmr_target']:,}원

{today.month}월 팀목표매출 : {log['month_target']:,}
현재달성매출 : {log['month_achieved']:,}원
현재달성율 : {pct}%"""

t1, t2, t3 = st.tabs(["출근보고", "15시보고", "마감보고"])
with t1:
    st.text_area("출근보고 (직접 수정 가능)", value=morning_default, height=180,
                 key=f"rt_morning_{rt_seq}", label_visibility="collapsed")
with t2:
    st.text_area("15시보고 (직접 수정 가능)", value=pm3_default, height=140,
                 key=f"rt_pm3_{rt_seq}", label_visibility="collapsed")
with t3:
    st.text_area("마감보고 (직접 수정 가능)", value=close_default, height=320,
                 key=f"rt_close_{rt_seq}", label_visibility="collapsed")
