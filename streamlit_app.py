import streamlit as st
import sqlite3, time, calendar as cal_lib, glob, os, re, threading, json, io
import pandas as pd
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

DB = "salesdb.db"
TT_DIR = "인트라넷 시간표"
FEE_FILE = "수강료.xlsx"
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

def _is_weekend_days(days):
    return bool(days) and set(days) <= {"토", "일"}

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
    subject = re.sub(r"/(주말|격일)$", "", m.group("subject")).strip()
    return dict(room=room, subject=subject, teacher=teacher, days=days, day_label=day_tok,
                start_date=m.group("start_date"), end_date=m.group("end_date"),
                cap=int(m.group("cap")), enrolled=int(m.group("enrolled")),
                assigned=int(m.group("assigned")), start_time=time_label)

@st.cache_data
def _load_timetable(_cache_key):
    sessions = []
    seen = set()  # 같은 강좌가 여러 스냅샷 파일에 중복 등록되는 걸 걸러냄
    for path in sorted(glob.glob(os.path.join(TT_DIR, "*.xls"))):
        try:
            df = pd.read_html(path, header=[0, 1])[0]
        except Exception:
            continue
        df = df[df.iloc[:, 0] != "정원"].reset_index(drop=True)
        times = df.iloc[:, 0].tolist()
        for col in df.columns[1:]:
            room = col[0]
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
                    dedup_key = (sess["subject"], sess["room"], sess["teacher"], sess["day_label"],
                                 sess["start_date"], sess["end_date"], sess["start_time"])
                    if dedup_key not in seen:
                        seen.add(dedup_key)
                        sessions.append(sess)
                i = j + 1
    return sessions

def get_timetable():
    files = glob.glob(os.path.join(TT_DIR, "*.xls"))
    key = tuple(sorted((os.path.basename(f), os.path.getmtime(f)) for f in files))
    return _load_timetable(key)

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

def sessions_active_on(d):
    ds = d.isoformat()
    out = [s for s in get_timetable() if s["start_date"] <= ds <= s["end_date"]]
    return sorted(out, key=lambda s: (s["start_time"], s["subject"]))

# ── 수강료표 파싱 ────────────────────────────────────────────
def _expand_fee_name(name):
    suffix = ""
    base = name
    sm = re.search(r"(\(방학\)|/주말)$", base)
    if sm:
        suffix = sm.group(1)
        base = base[:sm.start()]
    rm = re.match(r"^(.*?)(\d+)~(\d+)$", base)
    bases = [f"{rm.group(1)}{i}" for i in range(int(rm.group(2)), int(rm.group(3)) + 1)] if rm else [base]
    out = []
    for b in bases:
        parts = [p.strip() for p in b.split(",")] if "," in b else [b]
        out.extend(p + suffix for p in parts)
    return out

@st.cache_data
def _load_fees(_cache_key):
    fees = {}
    df = pd.read_excel(FEE_FILE, header=None)
    for _, row in df.iterrows():
        for name_col, fee_col in ((3, 4), (11, 12)):
            name, fee = row.get(name_col), row.get(fee_col)
            if isinstance(name, str) and isinstance(fee, (int, float)) and not pd.isna(fee):
                for key in _expand_fee_name(name.strip()):
                    fees[key] = int(fee)
    return fees

def lookup_fee(subject, weekend=False):
    if not os.path.exists(FEE_FILE):
        return None
    fees = _load_fees(os.path.getmtime(FEE_FILE))
    s = subject.strip()
    if s in fees:
        return fees[s]
    if weekend and f"{s}/주말" in fees:
        return fees[f"{s}/주말"]
    return None

# ── GitHub 기반 DB 영구 저장 (Streamlit Cloud 로컬 디스크는 슬립/재배포 시 초기화될 수 있음) ──
GITHUB_DB_PATH = "backup_salesdb.db"  # 로컬 salesdb.db와 별도 경로 (로컬 gitignore와 안 겹치게)

def _github_cfg():
    if not hasattr(st, "secrets"):
        return None, None
    token = st.secrets.get("GITHUB_TOKEN", "")
    if not token:
        return None, None
    return token, st.secrets.get("GITHUB_REPO", "songseonggi22-boop/list")

def restore_db_from_github():
    token, repo = _github_cfg()
    if not token or os.path.exists(DB):
        return
    import base64, urllib.request, json as _json
    url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DB_PATH}"
    try:
        req = urllib.request.Request(url, headers={"Authorization": f"token {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            content = base64.b64decode(_json.load(resp)["content"])
        with open(DB, "wb") as f:
            f.write(content)
    except Exception:
        pass  # 백업이 아직 없거나 네트워크 문제 → 새 DB로 시작

def backup_db_to_github():
    token, repo = _github_cfg()
    if not token:
        return
    import base64, urllib.request, json as _json
    url = f"https://api.github.com/repos/{repo}/contents/{GITHUB_DB_PATH}"
    headers = {"Authorization": f"token {token}"}
    sha = None
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            sha = _json.load(resp)["sha"]
    except Exception:
        pass
    try:
        with open(DB, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        body = {"message": "auto: DB 백업", "content": content_b64}
        if sha:
            body["sha"] = sha
        req = urllib.request.Request(url, data=_json.dumps(body).encode(),
                                      headers={**headers, "Content-Type": "application/json"}, method="PUT")
        urllib.request.urlopen(req, timeout=10)
    except Exception:
        pass  # ponytail: 네트워크 실패는 조용히 무시, 다음 쓰기 때 다시 시도됨

# ── DB ───────────────────────────────────────────────────────
@st.cache_resource
def get_db():
    restore_db_from_github()
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
        ctype TEXT DEFAULT '단과', assignee TEXT DEFAULT '',
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

    CREATE TABLE IF NOT EXISTS week_target(
        month TEXT, week_no INTEGER, assignee TEXT DEFAULT '',
        start_date TEXT, end_date TEXT, target INTEGER DEFAULT 0,
        PRIMARY KEY (month, week_no, assignee));
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
    try:
        c.execute("ALTER TABLE consultations ADD COLUMN assignee TEXT DEFAULT ''")
        c.commit()
    except sqlite3.OperationalError:
        pass  # 이미 있음
    try:
        c.execute("ALTER TABLE consultations ADD COLUMN result_status TEXT DEFAULT ''")
        c.commit()
    except sqlite3.OperationalError:
        pass  # 이미 있음
    try:
        c.execute("SELECT assignee FROM week_target LIMIT 1")
    except sqlite3.OperationalError:
        # 기존 week_target은 (month,week_no)만 PK라 assignee 컬럼을 못 넣음 → 재생성
        c.execute("ALTER TABLE week_target RENAME TO week_target_old")
        c.execute("""CREATE TABLE week_target(
            month TEXT, week_no INTEGER, assignee TEXT DEFAULT '',
            start_date TEXT, end_date TEXT, target INTEGER DEFAULT 0,
            PRIMARY KEY (month, week_no, assignee))""")
        c.execute("""INSERT INTO week_target(month,week_no,assignee,start_date,end_date,target)
                     SELECT month,week_no,'',start_date,end_date,target FROM week_target_old""")
        c.execute("DROP TABLE week_target_old")
        c.commit()
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
    backup_db_to_github()  # ponytail: 매 쓰기마다 동기 백업이라 약간 느려질 수 있음, 부담되면 디바운스 추가

def get_state(key, default=None):
    rows = q("SELECT value FROM app_state WHERE key=?", (key,))
    return rows[0][0] if rows else default

def set_state(key, value):
    run("INSERT INTO app_state(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value))

def get_team_members():
    try:
        return json.loads(get_state("team_members", "[]"))
    except Exception:
        return []

def add_team_member(name):
    members = get_team_members()
    if name and name not in members:
        members.append(name)
        set_state("team_members", json.dumps(members, ensure_ascii=False))

def remove_team_member(name):
    members = get_team_members()
    if name in members:
        members.remove(name)
        set_state("team_members", json.dumps(members, ensure_ascii=False))

def get_weekday_pct(person, day):
    return int(get_state(f"weekday_pct_{person}_{day}", "0") or 0)

def set_weekday_pct(person, day, pct):
    set_state(f"weekday_pct_{person}_{day}", str(int(pct)))

def get_tasks():
    rows = q("SELECT id,title,category,task_date,assignee,is_done FROM tasks ORDER BY is_done,created_at DESC")
    return [dict(zip("id title category task_date assignee is_done".split(), r)) for r in rows]

def get_consults():
    rows = q("""SELECT id,name,sched_date,sched_time,expected_revenue,ctype,assignee,result_status
                FROM consultations ORDER BY sched_date,sched_time""")
    return [dict(zip("id name sched_date sched_time expected_revenue ctype assignee result_status".split(), r))
            for r in rows]

def set_consult_status(cid, status):
    run("UPDATE consultations SET result_status=? WHERE id=?", (status, cid))

def add_task(title, cat, d, assignee):
    run("INSERT INTO tasks(title,category,task_date,assignee,is_done,created_at) VALUES(?,?,?,?,0,?)",
        (title, cat, d, assignee, int(time.time()*1000)))

def toggle_task(tid, v): run("UPDATE tasks SET is_done=? WHERE id=?", (int(v), tid))
def del_task(tid):       run("DELETE FROM tasks WHERE id=?", (tid,))

def add_consult(name, d, t, rev, ctype, assignee):
    run("INSERT INTO consultations(name,sched_date,sched_time,expected_revenue,ctype,assignee,created_at) VALUES(?,?,?,?,?,?,?)",
        (name, d, t, rev, ctype, assignee, int(time.time()*1000)))
def del_consult(cid): run("DELETE FROM consultations WHERE id=?", (cid,))

def update_consult(cid, name, d, t, rev, ctype, assignee):
    run("""UPDATE consultations SET name=?, sched_date=?, sched_time=?, expected_revenue=?, ctype=?, assignee=?
           WHERE id=?""", (name, d, t, rev, ctype, assignee, cid))

# ── 주차별 목표매출 (담당자별) ─────────────────────────────────
def get_week_targets(month, assignee):
    rows = q("SELECT week_no,start_date,end_date,target FROM week_target WHERE month=? AND assignee=? ORDER BY week_no",
             (month, assignee))
    return {r[0]: dict(start_date=r[1], end_date=r[2], target=r[3]) for r in rows}

def save_week_target(month, week_no, assignee, start_date, end_date, target):
    run("""INSERT INTO week_target(month,week_no,assignee,start_date,end_date,target) VALUES(?,?,?,?,?,?)
           ON CONFLICT(month,week_no,assignee) DO UPDATE SET
           start_date=excluded.start_date, end_date=excluded.end_date, target=excluded.target""",
        (month, week_no, assignee, start_date, end_date, target))

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

# ── 구글 시트 백업 (주기적 백업용, DB는 여전히 sqlite) ───────────
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

@st.cache_resource
def get_sheets_client():
    if not hasattr(st, "secrets") or "gcp_service_account" not in st.secrets:
        return None
    from google.oauth2.service_account import Credentials
    import gspread
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SHEETS_SCOPES)
    return gspread.authorize(creds)

def backup_to_sheets():
    import gspread
    client = get_sheets_client()
    sheet_id = st.secrets.get("SHEETS_SPREADSHEET_ID", "") if hasattr(st, "secrets") else ""
    if not client or not sheet_id:
        raise RuntimeError("구글 시트 백업 설정이 안 되어 있어요 (Secrets에 gcp_service_account / SHEETS_SPREADSHEET_ID 확인).")
    sh = client.open_by_key(sheet_id)

    def write_sheet(name, headers, rows):
        try:
            ws = sh.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=name, rows=max(100, len(rows) + 10), cols=max(10, len(headers)))
        ws.clear()
        ws.update([headers] + [[str(v) for v in row] for row in rows])

    tasks_ = get_tasks()
    write_sheet("tasks", ["id", "title", "category", "task_date", "assignee", "is_done"],
                [[t["id"], t["title"], t["category"], t["task_date"], t["assignee"], t["is_done"]] for t in tasks_])

    consults_ = get_consults()
    write_sheet("consultations", ["id", "name", "sched_date", "sched_time", "expected_revenue", "ctype"],
                [[c["id"], c["name"], c["sched_date"], c["sched_time"], c["expected_revenue"], c["ctype"]] for c in consults_])

    logs_ = q(f"SELECT {','.join(LOG_COLS)} FROM daily_log ORDER BY log_date")
    write_sheet("daily_log", LOG_COLS, [list(r) for r in logs_])

# ── 배정 문구 생성기 ──────────────────────────────────────────
def fdate(iso):
    if not iso: return ""
    p = iso.split("-")
    return f"{int(p[1])}.{int(p[2])}"

def gen_text(type_, **k):
    nd, nt = fdate(k.get("nd","")), k.get("nt","")
    ns, nts = k.get("ns",""), k.get("nts","")
    ns_full = f"{ns}/{nts}"
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

# 업무일 기준 이번 주 월요일 → 주중업무 게이지는 월요일이 되면 새로 시작
_workday_date = date.fromisoformat(workday)
week_start = (_workday_date - timedelta(days=_workday_date.weekday())).isoformat()

# 하루 한 번 자동 구글시트 백업 (설정 안 돼있으면 조용히 건너뜀)
if get_state("last_sheets_backup") != today_str and get_sheets_client():
    try:
        backup_to_sheets()
        set_state("last_sheets_backup", today_str)
    except Exception:
        pass  # 다음 로드 때 다시 시도

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
            cassignee = st.text_input("담당자", placeholder="담당자")
            if st.form_submit_button("저장", use_container_width=True) and cname.strip():
                add_consult(cname.strip(), cdate.isoformat(), ctime, int(crev), ctype, cassignee.strip()); st.rerun()

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
            if st.session_state.get("edit_consult_id") == c["id"]:
                with st.form(f"editcf_{c['id']}"):
                    e1, e2 = st.columns(2)
                    ename = e1.text_input("이름", value=c["name"])
                    edate = e2.date_input("날짜", value=datetime.strptime(c["sched_date"], "%Y-%m-%d").date())
                    e3, e4, e5 = st.columns(3)
                    etime = e3.text_input("시간", value=c["sched_time"])
                    erev  = e4.number_input("예정매출(원)", min_value=0, step=10000, value=c["expected_revenue"])
                    ectype = e5.selectbox("구분", ["단과", "정규"], index=0 if c["ctype"] == "단과" else 1)
                    eassignee = st.text_input("담당자", value=c["assignee"])
                    s1, s2 = st.columns(2)
                    if s1.form_submit_button("저장", use_container_width=True, key=f"save_ec_{c['id']}"):
                        update_consult(c["id"], ename.strip(), edate.isoformat(), etime,
                                       int(erev), ectype, eassignee.strip())
                        st.session_state["edit_consult_id"] = None
                        st.rerun()
                    if s2.form_submit_button("취소", use_container_width=True, key=f"cancel_ec_{c['id']}"):
                        st.session_state["edit_consult_id"] = None
                        st.rerun()
            else:
                cc1, cc2, cc3 = st.columns([1, 0.13, 0.13])
                cc1.markdown(f"""
<div style='background:#eef2ff;border-left:3px solid #4f46e5;border-radius:6px;
            padding:7px 12px;margin:3px 0;display:flex;justify-content:space-between;align-items:center'>
  <span style='font-size:12px;font-weight:700;color:#3730a3'>{c['name']} <span style='font-weight:500;color:#8b8bd8'>({c['ctype']})</span></span>
  <span style='font-size:11px;color:#6366f1'>{c['sched_time']}</span>
  <span style='font-size:11px;font-weight:700;color:#4f46e5'>{c['expected_revenue']:,}원</span>
</div>""", unsafe_allow_html=True)
                if cc2.button("✏️", key=f"ec{c['id']}"):
                    st.session_state["edit_consult_id"] = c["id"]
                    st.rerun()
                if cc3.button("✕", key=f"dc{c['id']}"):
                    del_consult(c["id"]); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="db-card"><div class="db-card-title">☑ TO DO LIST</div>', unsafe_allow_html=True)
    st.caption("매일 한국시간 06:00에 완료 체크가 초기화돼요.")

    with st.expander("👥 담당자 관리"):
        mc1, mc2 = st.columns([3, 1])
        new_member = mc1.text_input("담당자 이름 추가", key="new_member_name", label_visibility="collapsed",
                                     placeholder="담당자 이름 입력")
        if mc2.button("추가", key="add_member_btn", use_container_width=True) and new_member.strip():
            add_team_member(new_member.strip()); st.rerun()
        for m in get_team_members():
            rc1, rc2 = st.columns([3, 1])
            rc1.caption(m)
            if rc2.button("삭제", key=f"del_member_{m}"):
                remove_team_member(m); st.rerun()

    CATS = [
        ("우선순위1","우선순위 1","kp1"),
        ("우선순위2","우선순위 2","kp2"),
        ("주중업무", "주중업무",  "kp3"),
    ]
    groups = get_team_members() + ["미지정"]
    todo_tabs = st.tabs([f"👤 {p}" for p in groups])

    for tab, person in zip(todo_tabs, groups):
      with tab:
        person_tasks = [t for t in tasks if (t["assignee"] or "미지정") == person]
        k1, k2, k3 = st.columns(3, gap="small")

        for col_w, (cat_key, cat_label, kp_cls) in zip([k1, k2, k3], CATS):
            with col_w:
                st.markdown(f'<div class="k-head {kp_cls}">{cat_label}</div>', unsafe_allow_html=True)

                if cat_key == "주중업무":
                    pct = get_weekday_pct(person, week_start)
                    st.markdown(f"""
<div style='margin-bottom:8px'>
  <div style='display:flex;justify-content:space-between;font-size:11px;color:#666'>
    <span>진행률</span><span><b>{pct}%</b></span>
  </div>
  <div style='background:#eee;border-radius:6px;height:8px;overflow:hidden'>
    <div style='background:#2f80ed;height:100%;width:{pct}%'></div>
  </div>
</div>""", unsafe_allow_html=True)
                    new_pct = st.slider("진행률", 0, 100, pct, step=5,
                                        key=f"pct_slider_{person}", label_visibility="collapsed")
                    if new_pct != pct:
                        set_weekday_pct(person, week_start, new_pct); st.rerun()

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

                # 할 일 추가 폼 (이 탭 담당자로 자동 배정, 필요하면 수정 가능)
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

# ── 주차별 목표매출 (담당자별 탭) ──────────────────────────────
st.markdown("""
<div class="db-card">
  <div class="db-card-title">🎯 목표매출 설정</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    월 전체 목표와 1~4주차 목표를 담당자별로 따로 설정하고 진행률을 확인합니다.
  </div>
</div>""", unsafe_allow_html=True)

target_month_pick = st.date_input("기준 월", value=today, key="target_month_pick")
target_month = target_month_pick.strftime("%Y-%m")
_, days_in_month = cal_lib.monthrange(target_month_pick.year, target_month_pick.month)
DEFAULT_WEEK_RANGES = [(1, 7), (8, 14), (15, 21), (22, days_in_month)]

def target_card_html(title, period_label, target, actual, highlight=False):
    pct = round(actual / target * 100, 1) if target else 0.0
    bar_pct = min(pct, 100)
    if highlight:
        bg, tx, sub, line, track = "#1e3a5f", "#fff", "#9db4cf", "rgba(255,255,255,0.25)", "rgba(255,255,255,0.15)"
    else:
        bg, tx, sub, line, track = "#fff", "#333", "#999", "#eee", "#eee"
    return f"""
<div style="background:{bg};border-radius:14px;padding:18px 12px;text-align:center;
            box-shadow:0 2px 8px rgba(0,0,0,0.05);height:100%">
  <div style="font-size:10px;color:{sub}">기간: ({period_label})</div>
  <div style="font-size:14px;font-weight:700;margin:6px 0;color:{tx}">{title}</div>
  <div style="font-size:20px;font-weight:800;color:{tx}">{target:,}</div>
  <hr style="margin:10px 0;border-color:{line}">
  <div style="font-size:11px;color:{sub}">매출현황</div>
  <div style="font-size:17px;font-weight:700;color:{tx};margin:4px 0">{actual:,}</div>
  <div style="font-size:12px;color:{sub}">달성율 <b style="color:#2f80ed">{pct}%</b></div>
  <div style="background:{track};border-radius:6px;height:6px;margin-top:8px;overflow:hidden">
    <div style="background:#2f80ed;height:100%;width:{bar_pct}%"></div>
  </div>
</div>"""

def render_target_dashboard(scope_key):
    scope_consults = consults if scope_key == "전체" else [c for c in consults if c["assignee"] == scope_key]
    wt_saved = get_week_targets(target_month, scope_key)
    total_target_saved = int(get_state(f"month_total_target_{target_month}_{scope_key}", "0") or 0)

    with st.expander("✏️ 목표 설정/수정", expanded=not wt_saved):
        with st.form(f"wtf_{scope_key}_{target_month}"):
            total_target = st.number_input(f"{target_month} 총 목표매출(원)", min_value=0, step=100000,
                                            value=total_target_saved, key=f"tt_{scope_key}")
            week_inputs = []
            for wn in range(1, 5):
                default = wt_saved.get(wn)
                d_start, d_end = DEFAULT_WEEK_RANGES[wn - 1]
                d_end = min(d_end, days_in_month)
                sd_default = (datetime.strptime(default["start_date"], "%Y-%m-%d").date()
                              if default else target_month_pick.replace(day=d_start))
                ed_default = (datetime.strptime(default["end_date"], "%Y-%m-%d").date()
                              if default else target_month_pick.replace(day=d_end))
                tg_default = default["target"] if default else 0
                c1, c2, c3 = st.columns(3)
                sd = c1.date_input(f"{wn}주차 시작", value=sd_default, key=f"wt_sd_{scope_key}_{wn}")
                ed = c2.date_input(f"{wn}주차 종료", value=ed_default, key=f"wt_ed_{scope_key}_{wn}")
                tg = c3.number_input(f"{wn}주차 목표(원)", min_value=0, step=100000,
                                      value=tg_default, key=f"wt_tg_{scope_key}_{wn}")
                week_inputs.append((wn, sd, ed, tg))
            if st.form_submit_button("저장", use_container_width=True):
                set_state(f"month_total_target_{target_month}_{scope_key}", str(int(total_target)))
                for wn, sd, ed, tg in week_inputs:
                    save_week_target(target_month, wn, scope_key, sd.isoformat(), ed.isoformat(), int(tg))
                st.rerun()

    cols = st.columns(5)
    for wn in range(1, 5):
        info = wt_saved.get(wn)
        with cols[wn - 1]:
            if info:
                wk_actual = sum(c["expected_revenue"] for c in scope_consults
                                 if info["start_date"] <= c["sched_date"] <= info["end_date"])
                st.markdown(target_card_html(f"{wn}주차 목표", f"{info['start_date']}~{info['end_date']}",
                                              info["target"], wk_actual), unsafe_allow_html=True)
            else:
                st.markdown(target_card_html(f"{wn}주차 목표", "미설정", 0, 0), unsafe_allow_html=True)
    with cols[4]:
        starts = [wt_saved[w]["start_date"] for w in wt_saved]
        ends = [wt_saved[w]["end_date"] for w in wt_saved]
        period_label = f"{min(starts)}~{max(ends)}" if starts else target_month
        month_actual = sum(c["expected_revenue"] for c in scope_consults if c["sched_date"][:7] == target_month)
        st.markdown(target_card_html("총목표매출", period_label, total_target_saved, month_actual, highlight=True),
                    unsafe_allow_html=True)

assignee_pool = get_team_members()
target_tabs = st.tabs(["전체"] + assignee_pool)
for tab, scope_key in zip(target_tabs, ["전체"] + assignee_pool):
    with tab:
        render_target_dashboard(scope_key)

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 담당자별 실적 집계 ─────────────────────────────────────────
st.markdown("""
<div class="db-card">
  <div class="db-card-title">📊 담당자별 실적</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    상담 일정에 등록된 담당자 기준으로 집계합니다.
  </div>
</div>""", unsafe_allow_html=True)

period = st.radio("기간", ["오늘", "이번 주", "이번 달", "전체"], horizontal=True, label_visibility="collapsed")
if period == "오늘":
    period_consults = [c for c in consults if c["sched_date"] == today_str]
elif period == "이번 주":
    week_start = (today - timedelta(days=today.weekday())).isoformat()
    period_consults = [c for c in consults if week_start <= c["sched_date"] <= today_str]
elif period == "이번 달":
    month_start = today.replace(day=1).isoformat()
    period_consults = [c for c in consults if month_start <= c["sched_date"] <= today_str]
else:
    period_consults = consults

perf = {}
for c in period_consults:
    who = c["assignee"] or "미지정"
    p = perf.setdefault(who, {"건수": 0, "예정매출": 0, "정규": 0, "단과": 0})
    p["건수"] += 1
    p["예정매출"] += c["expected_revenue"]
    p[c["ctype"]] = p.get(c["ctype"], 0) + 1

if perf:
    perf_df = pd.DataFrame([{"담당자": k, **v} for k, v in perf.items()]).sort_values("예정매출", ascending=False)
    st.dataframe(perf_df, use_container_width=True, hide_index=True)
else:
    st.caption("해당 기간에 상담 데이터가 없어요.")

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

with st.expander("🔍 날짜 클릭 → 그날 진행 중인 강좌 선택 (신규 배정용)", expanded=False):
    pick_date = st.date_input("날짜", value=today, key="tt_pick_date")
    day_sessions = sessions_active_on(pick_date)
    if day_sessions:
        by_time = {}
        for s in day_sessions:
            by_time.setdefault(s["start_time"], []).append(s)
        si = 0
        for t in sorted(by_time):
            st.markdown(f"""<div style='font-size:12px;font-weight:700;color:#4f46e5;margin:8px 0 4px'>
                        🕐 {t}</div>""", unsafe_allow_html=True)
            for s in sorted(by_time[t], key=lambda x: x["subject"]):
                weekend = _is_weekend_days(s["days"])
                label = f"{s['subject']} ({s['room']}, {s['teacher']}) 개강 {s['start_date']}" + ("  [주말]" if weekend else "")
                key = f"ttpick_{si}_{cb_seq}"
                st.checkbox(label, key=key)
                candidates.append(dict(key=key, nd=s["start_date"],
                                        nt=s["start_time"], ns=s["subject"], nts="주말" if weekend else ""))
                si += 1
    else:
        st.caption("이 날짜에 진행 중인 강좌가 없어요.")

with st.expander("🔎 과목명으로 검색 (신규 배정용)"):
    search_q = st.text_input("과목명/강사명 검색", key="tt_search_q")
    if search_q:
        hits = [s for s in get_timetable()
                if search_q.strip() in s["subject"] or search_q.strip() in s["teacher"]]
        hits = sorted(hits, key=lambda s: (s["subject"], s["start_time"]))
        if hits:
            for i, s in enumerate(hits):
                weekend = _is_weekend_days(s["days"])
                label = (f"{s['subject']} — {s['start_time']} ({s['room']}, {s['teacher']}) "
                         f"개강 {s['start_date']}~종강 {s['end_date']}") + ("  [주말]" if weekend else "")
                key = f"ttsearch_{i}_{cb_seq}"
                st.checkbox(label, key=key)
                candidates.append(dict(key=key, nd=s["start_date"],
                                        nt=s["start_time"], ns=s["subject"], nts="주말" if weekend else ""))
        else:
            st.caption("일치하는 강좌가 없어요.")

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
                weekend = _is_weekend_days(c["days"])
                label = f"{c['subject']} — {c['start_time']} 개강 {c['start_date']}" + ("  [주말]" if weekend else "")
                key = f"pt_pick_{i}_{cb_seq}"
                st.checkbox(label, key=key)
                candidates.append(dict(key=key, nd=c["start_date"], nt=c["start_time"],
                                        ns=c["subject"], nts="주말" if weekend else ""))
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
            weekend = _is_weekend_days(_expand_days(c.get("day_label", "")))
            label = f"{c.get('subject','')} — {c.get('start_time','')} 개강 {c.get('start_date','')}" + ("  [주말]" if weekend else "")
            key = f"img_pick_{i}_{cb_seq}"
            st.checkbox(label, key=key)
            candidates.append(dict(key=key, nd=c.get("start_date", ""), nt=c.get("start_time", ""),
                                    ns=c.get("subject", ""), nts="주말" if weekend else ""))
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

def _course_picker(label, key):
    c1, c2 = st.columns([1, 2])
    mode = c1.radio("찾기", ["날짜로", "검색으로"], key=f"{key}_mode", horizontal=True, label_visibility="collapsed")
    if mode == "날짜로":
        d = c2.date_input(label, value=today, key=f"{key}_date", label_visibility="collapsed")
        opts = sessions_active_on(d)
    else:
        q = c2.text_input(label, key=f"{key}_q", placeholder=f"{label} 과목명 검색", label_visibility="collapsed")
        opts = sorted([s for s in get_timetable() if q.strip() and q.strip() in s["subject"]],
                      key=lambda s: (s["subject"], s["start_time"])) if q else []
    if not opts:
        st.caption("일치하는 강좌가 없어요.")
        return None
    labels = [f"{s['subject']} — {s['start_time']} ({s['room']}) 개강 {s['start_date']}" for s in opts]
    idx = st.selectbox(f"{label} 강좌", range(len(opts)), format_func=lambda i: labels[i],
                       key=f"{key}_sel", label_visibility="collapsed")
    return opts[idx]

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
    st.markdown("<div style='font-size:11px;color:#999;margin-bottom:4px'>▸ 이전 강좌</div>", unsafe_allow_html=True)
    old_s = _course_picker("이전", "cg_old")
    st.markdown("<div style='font-size:11px;color:#999;margin:8px 0 4px'>▸ 변경 후 강좌</div>", unsafe_allow_html=True)
    new_s = _course_picker("변경 후", "cg_new")
    if old_s and new_s:
        auto_ofee = lookup_fee(old_s["subject"], _is_weekend_days(old_s["days"])) or 0
        auto_nfee = lookup_fee(new_s["subject"], _is_weekend_days(new_s["days"])) or 0
        c1, c2 = st.columns(2)
        ofee = c1.number_input(f"이전 수강료 (자동조회: {auto_ofee:,}원)", min_value=0, step=10000,
                                value=auto_ofee, key="cg_ofee")
        nfee = c2.number_input(f"변경 후 수강료 (자동조회: {auto_nfee:,}원)", min_value=0, step=10000,
                                value=auto_nfee, key="cg_nfee")
        if st.button("✨ 문구 생성", use_container_width=True, key="cg_gen"):
            result = gen_text("과목변경", os_=old_s["subject"], ofee=int(ofee),
                              nd=new_s["start_date"], nt=new_s["start_time"], ns=new_s["subject"],
                              nts="주말" if _is_weekend_days(new_s["days"]) else "", nfee=int(nfee))

elif gtype == "배정 취소":
    cancel_s = _course_picker("취소할", "cx")
    if cancel_s and st.button("✨ 문구 생성", use_container_width=True, key="cx_gen"):
        result = gen_text("취소", od=cancel_s["start_date"], ot=cancel_s["start_time"], os_=cancel_s["subject"])

elif gtype == "날짜변경 배정":
    st.markdown("<div style='font-size:11px;color:#999;margin-bottom:4px'>▸ 이전 일정</div>", unsafe_allow_html=True)
    old_s = _course_picker("이전", "dc_old")
    st.markdown("<div style='font-size:11px;color:#999;margin:8px 0 4px'>▸ 변경 후 일정</div>", unsafe_allow_html=True)
    new_s = _course_picker("변경 후", "dc_new")
    if old_s and new_s and st.button("✨ 문구 생성", use_container_width=True, key="dc_gen"):
        result = gen_text("날짜변경", od=old_s["start_date"], ot=old_s["start_time"], os_=old_s["subject"],
                          nd=new_s["start_date"], nt=new_s["start_time"])

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

STATUS_OPTS = ["미정", "등록", "COD", "미등록"]
today_consults = [c for c in consults if c["sched_date"] == today_str]

st.markdown("<div style='font-size:13px;font-weight:700;margin-bottom:6px'>오늘 상담 결과 체크</div>",
            unsafe_allow_html=True)
if today_consults:
    for c in today_consults:
        rc1, rc2 = st.columns([2, 2])
        rc1.caption(f"{c['name']} · {c['sched_time']} · {c['expected_revenue']:,}원")
        cur = c["result_status"] if c["result_status"] in STATUS_OPTS else "미정"
        new_status = rc2.radio("결과", STATUS_OPTS, index=STATUS_OPTS.index(cur), horizontal=True,
                                key=f"cstat_{c['id']}", label_visibility="collapsed")
        if new_status != cur:
            set_consult_status(c["id"], new_status); st.rerun()
else:
    st.caption("오늘 등록된 상담이 없어요.")

registered_ct = sum(1 for c in today_consults if c["result_status"] == "등록")
cod_ct = sum(1 for c in today_consults if c["result_status"] == "COD")
unregistered_ct = sum(1 for c in today_consults if c["result_status"] == "미등록")
auto_revenue = sum(c["expected_revenue"] for c in today_consults if c["result_status"] == "등록")

rev_c1, rev_c2 = st.columns(2)
in_revenue = rev_c1.number_input(f"금일매출결과 (자동조회: {auto_revenue:,}원)", min_value=0, step=10000,
                                  value=log["actual_revenue"] or auto_revenue, key="in_actual_revenue")
in_refund = rev_c2.number_input("환불", min_value=0, step=10000, value=log["refund"], key="in_refund")
if st.button("💾 금일매출결과·환불 저장"):
    run("UPDATE daily_log SET actual_revenue=?, refund=? WHERE log_date=?",
        (int(in_revenue), int(in_refund), today_str))
    st.rerun()

this_month = today.strftime("%Y-%m")
month_target_live = int(get_state(f"month_total_target_{this_month}_전체", "0") or 0)
month_achieved_live = sum(c["expected_revenue"] for c in consults if c["sched_date"][:7] == this_month)
pct = round(month_achieved_live / month_target_live * 100) if month_target_live else 0

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
 등록 : {registered_ct}
 COD : {cod_ct}
 미등록 : {unregistered_ct}

금일매출결과 :{int(in_revenue):,}원
환불 :{int(in_refund):,}원

통화시간
{log['rep1_name']} {log['rep1_call']} (상담 {today_cnt}건)
{log['rep2_name']} {log['rep2_call']}

익일예정상담 : {tmr_cnt}건
익일예정매출 : {tmr_rev:,}원
익일목표매출 : {log['tmr_target']:,}원

{today.month}월 팀목표매출 : {month_target_live:,}
현재달성매출 : {month_achieved_live:,}원
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

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 구글 시트 백업 ────────────────────────────────────────────
st.markdown("""
<div class="db-card">
  <div class="db-card-title">☁️ 구글 시트 백업</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    하루 한 번 자동으로 상담·할일·일지 데이터를 백업합니다. DB는 그대로 sqlite를 씁니다.
  </div>
</div>""", unsafe_allow_html=True)

if not get_sheets_client():
    st.caption("아직 설정 안 됨 — Streamlit Secrets에 gcp_service_account / SHEETS_SPREADSHEET_ID를 추가하세요.")
else:
    last_backup = get_state("last_sheets_backup", "없음")
    st.caption(f"마지막 백업: {last_backup}")
    if st.button("지금 백업하기"):
        try:
            backup_to_sheets()
            set_state("last_sheets_backup", today_str)
            st.success("백업 완료!")
        except Exception as e:
            st.error(f"백업 실패: {e}")
