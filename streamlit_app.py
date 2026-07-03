import streamlit as st
import sqlite3, time, calendar as cal_lib
from datetime import date, datetime

DB = "salesdb.db"

DEFAULT_FIXED = [
    "아침 루틴 확인", "미확인 문자 회신",
    "전화 확인", "일일 보고 작성", "전달사항 확인",
]

# ── DB ───────────────────────────────────────────────────────
@st.cache_resource
def get_db():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.executescript("""
    CREATE TABLE IF NOT EXISTS fixed_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, order_num INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS daily_checks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fixed_id INTEGER NOT NULL, check_date TEXT NOT NULL,
        UNIQUE(fixed_id, check_date));

    CREATE TABLE IF NOT EXISTS adhoc_tasks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, is_done INTEGER DEFAULT 0,
        task_date TEXT NOT NULL, created_at INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS consultations(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, sched_date TEXT NOT NULL,
        sched_time TEXT DEFAULT '', expected_revenue INTEGER DEFAULT 0,
        note TEXT DEFAULT '', created_at INTEGER DEFAULT 0);
    """)
    if c.execute("SELECT COUNT(*) FROM fixed_items").fetchone()[0] == 0:
        for i, t in enumerate(DEFAULT_FIXED):
            c.execute("INSERT INTO fixed_items(title,order_num) VALUES(?,?)", (t, i))
    c.commit()
    return c

def q(sql, args=()):  return get_db().execute(sql, args).fetchall()
def run(sql, args=()):get_db().execute(sql, args); get_db().commit()

def is_checked(fid, d): return bool(q("SELECT 1 FROM daily_checks WHERE fixed_id=? AND check_date=?",(fid,d)))
def toggle_fix(fid, d, on):
    if on: run("INSERT OR IGNORE INTO daily_checks(fixed_id,check_date) VALUES(?,?)",(fid,d))
    else:  run("DELETE FROM daily_checks WHERE fixed_id=? AND check_date=?",(fid,d))

def get_fixed():  return q("SELECT id,title FROM fixed_items ORDER BY order_num")
def get_adhoc(d): return q("SELECT id,title,is_done FROM adhoc_tasks WHERE task_date=? ORDER BY created_at",(d,))
def add_adhoc(t,d): run("INSERT INTO adhoc_tasks(title,is_done,task_date,created_at) VALUES(?,0,?,?)",(t,d,int(time.time()*1000)))
def toggle_adhoc(tid,v): run("UPDATE adhoc_tasks SET is_done=? WHERE id=?",(int(v),tid))
def del_adhoc(tid):      run("DELETE FROM adhoc_tasks WHERE id=?",(tid,))

def get_consults():
    rows = q("SELECT id,name,sched_date,sched_time,expected_revenue,note FROM consultations ORDER BY sched_date,sched_time")
    return [dict(zip("id name sched_date sched_time expected_revenue note".split(),r)) for r in rows]
def add_consult(name,d,t,rev,note): run("INSERT INTO consultations(name,sched_date,sched_time,expected_revenue,note,created_at) VALUES(?,?,?,?,?,?)",(name,d,t,rev,note,int(time.time()*1000)))
def del_consult(cid): run("DELETE FROM consultations WHERE id=?",(cid,))

# ── 배정 문구 생성기 ──────────────────────────────────────────
def fdate(iso):
    if not iso: return ""
    p = iso.split("-")
    return f"{int(p[1])}.{int(p[2])}"

def gen_text(type_, **k):
    nd, nt = fdate(k.get("nd","")), k.get("nt","")
    ns = k.get("ns","")
    nts = k.get("nts","")
    ns_full = f"{ns}/{nts}" if nts else ns
    nfee = k.get("nfee", 0)
    od, ot = fdate(k.get("od","")), k.get("ot","")
    os_ = k.get("os_","")
    ofee = k.get("ofee", 0)

    if type_ == "신규":
        return f"미배정 -> {nd} {nt} {ns_full} 배정"

    if type_ == "과목변경":
        if ofee or nfee:
            diff = abs(nfee - ofee)
            diff_txt = f"\n[차액 {diff:,}원 무시]" if diff > 0 else ""
            return (f"미배정 {os_} (수강료:{ofee:,}원) -> {nd} {nt} {ns_full} "
                    f"(수강료:{nfee:,}원) 과목변경 배정{diff_txt}")
        return f"미배정 {os_} -> {nd} {nt} {ns_full} 과목변경 배정"

    if type_ == "취소":
        return f"{od} {ot} {os_} 배정 -> 미배정"

    if type_ == "날짜변경":
        return f"{od} {ot} {os_} 배정 -> {nd} {nt} {os_} 배정"

# ── 앱 설정 ───────────────────────────────────────────────────
st.set_page_config(page_title="영업 대시보드", page_icon="📊",
                   layout="wide", initial_sidebar_state="collapsed")

ss = st.session_state
if "assign_out" not in ss: ss.assign_out = ""

today     = date.today()
today_str = today.isoformat()
now_str   = datetime.now().strftime("%H:%M")
wd        = ["월","화","수","목","금","토","일"][today.weekday()]

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""<style>
#MainMenu,footer,[data-testid="stHeader"],[data-testid="stToolbar"]{display:none!important}
.block-container{padding:1.6rem 2rem 3rem!important;max-width:100%!important}
[data-testid="stAppViewContainer"],.main{background:#F1F5F9!important}

[data-testid="stButton"]>button{
  padding:5px 14px!important;font-size:12px!important;border-radius:7px!important;
  font-weight:500!important;min-height:0!important;line-height:1.5!important;
  border:1.5px solid #E2E8F0!important;background:white!important;color:#475569!important}
[data-testid="stButton"]>button:hover{background:#F8FAFC!important;color:#1E293B!important}
[data-testid="stForm"]{border:none!important;padding:0!important}
.stTextInput input,.stTextArea textarea{
  border-radius:8px!important;border:1.5px solid #E2E8F0!important;
  font-size:13px!important;background:white!important}
.stSelectbox>div>div,.stDateInput input,.stNumberInput input{
  border-radius:8px!important;border:1.5px solid #E2E8F0!important;font-size:13px!important}
.stCheckbox label{font-size:13px!important;color:#374151!important}
div[data-testid="stHorizontalBlock"]{gap:1rem!important}
[data-testid="stExpander"]{border:1.5px solid #E2E8F0!important;border-radius:10px!important;background:white!important}
hr{margin:.6rem 0!important;border-color:#F1F5F9!important}
</style>""", unsafe_allow_html=True)

# ── 데이터 ───────────────────────────────────────────────────
fixed_items = get_fixed()
adhoc_tasks = get_adhoc(today_str)
consults    = get_consults()
today_c     = [c for c in consults if c["sched_date"] == today_str]
today_rev   = sum(c["expected_revenue"] for c in today_c)
total_rev   = sum(c["expected_revenue"] for c in consults)
fix_done    = sum(1 for fi in fixed_items if is_checked(fi[0], today_str))
ad_done     = sum(1 for t in adhoc_tasks if t[2])
chk_total   = len(fixed_items) + len(adhoc_tasks)
chk_done    = fix_done + ad_done

# ── 헤더 ─────────────────────────────────────────────────────
h1, h2 = st.columns([3, 1])
h1.markdown(f"""
<div style='padding:2px 0 18px'>
  <div style='font-size:26px;font-weight:800;color:#0F172A;letter-spacing:-.5px'>📊 영업 대시보드</div>
  <div style='font-size:13px;color:#94A3B8;margin-top:4px;font-weight:500'>
    {today.strftime('%Y년 %m월 %d일')} ({wd}요일)
  </div>
</div>""", unsafe_allow_html=True)
h2.markdown(f"""
<div style='text-align:right;padding-top:4px'>
  <div style='font-size:46px;font-weight:200;color:#0F172A;letter-spacing:6px;font-family:monospace'>{now_str}</div>
</div>""", unsafe_allow_html=True)

# ── 요약 카드 4개 ─────────────────────────────────────────────
for col, (label, val, color, emo) in zip(st.columns(4), [
    ("오늘 상담",    f"{len(today_c)}건",      "#4F46E5", "📅"),
    ("오늘 예정매출", f"{today_rev:,}원",       "#059669", "💰"),
    ("업무 완료",    f"{chk_done}/{chk_total}", "#D97706", "✅"),
    ("전체 예정매출", f"{total_rev:,}원",       "#DC2626", "📈"),
]):
    col.markdown(f"""
<div style='background:white;border-radius:12px;padding:16px 20px;
            border:1px solid #E2E8F0;border-top:3px solid {color}'>
  <div style='font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.5px'>{emo} {label}</div>
  <div style='font-size:28px;font-weight:800;color:#0F172A;margin-top:5px'>{val}</div>
</div>""", unsafe_allow_html=True)

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 메인 2컬럼 ────────────────────────────────────────────────
left, right = st.columns([1, 1.5], gap="medium")

# ─── 왼쪽: 체크리스트 ─────────────────────────────────────────
with left:
    st.markdown("""<div style='background:white;border-radius:12px;border:1px solid #E2E8F0;
    padding:18px 18px 6px'><div style='font-size:14px;font-weight:700;color:#0F172A;
    margin-bottom:12px'>✅ 업무 체크리스트</div></div>""", unsafe_allow_html=True)

    # 고정 항목
    for fid, ftitle in fixed_items:
        chk = is_checked(fid, today_str)
        c1, c2 = st.columns([0.1, 1])
        new = c1.checkbox("", value=chk, key=f"f{fid}")
        if new != chk:
            toggle_fix(fid, today_str, new); st.rerun()
        sty = "text-decoration:line-through;color:#CBD5E1" if chk else "color:#374151"
        c2.markdown(f"<div style='font-size:13px;{sty};padding-top:6px'>{ftitle}</div>",
                    unsafe_allow_html=True)

    st.divider()

    # 오늘 추가 항목
    st.markdown("<div style='font-size:11px;font-weight:700;color:#94A3B8;margin-bottom:6px'>오늘 추가 할 일</div>",
                unsafe_allow_html=True)
    for tid, ttitle, tdone in adhoc_tasks:
        c1, c2, c3 = st.columns([0.1, 1, 0.1])
        new = c1.checkbox("", value=bool(tdone), key=f"a{tid}")
        if new != bool(tdone):
            toggle_adhoc(tid, new); st.rerun()
        sty = "text-decoration:line-through;color:#CBD5E1" if tdone else "color:#374151"
        c2.markdown(f"<div style='font-size:13px;{sty};padding-top:6px'>{ttitle}</div>",
                    unsafe_allow_html=True)
        if c3.button("✕", key=f"ad{tid}"):
            del_adhoc(tid); st.rerun()

    with st.form("adhoc_f", clear_on_submit=True):
        nc = st.columns([1, 0.28])
        newtask = nc[0].text_input("", placeholder="할 일 추가...", label_visibility="collapsed")
        if nc[1].form_submit_button("추가", use_container_width=True) and newtask.strip():
            add_adhoc(newtask.strip(), today_str); st.rerun()
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# ─── 오른쪽: 상담 일정 ─────────────────────────────────────────
with right:
    st.markdown("""<div style='background:white;border-radius:12px;border:1px solid #E2E8F0;
    padding:18px 18px 6px'><div style='font-size:14px;font-weight:700;color:#0F172A;
    margin-bottom:8px'>📅 상담 일정</div></div>""", unsafe_allow_html=True)

    # 오늘 상담 하이라이트
    if today_c:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:#4F46E5;margin:6px 0 6px'>🔵 오늘 상담 {len(today_c)}건 · 예정 {today_rev:,}원</div>",
                    unsafe_allow_html=True)
        for c in today_c:
            st.markdown(f"""
<div style='background:#EEF2FF;border-left:3px solid #4F46E5;border-radius:6px;
            padding:8px 14px;margin-bottom:5px;display:flex;justify-content:space-between'>
  <span style='font-size:13px;font-weight:700;color:#3730A3'>{c['name']}</span>
  <span style='font-size:12px;color:#6366F1'>{c['sched_time']}</span>
  <span style='font-size:12px;font-weight:600;color:#4F46E5'>{c['expected_revenue']:,}원</span>
</div>""", unsafe_allow_html=True)

    st.divider()

    # 전체 목록
    if consults:
        st.markdown("<div style='font-size:11px;font-weight:700;color:#94A3B8;margin-bottom:8px'>전체 상담 목록</div>",
                    unsafe_allow_html=True)
        for c in consults:
            is_td = c["sched_date"] == today_str
            is_past = c["sched_date"] < today_str
            dc = "#4F46E5" if is_td else ("#94A3B8" if is_past else "#374151")
            cc = st.columns([1.4, 0.9, 1.1, 0.15])
            cc[0].markdown(f"<div style='font-size:13px;font-weight:600;color:#0F172A;padding-top:4px'>{c['name']}</div>",
                           unsafe_allow_html=True)
            cc[1].markdown(f"<div style='font-size:12px;color:{dc};padding-top:4px'>{c['sched_date'][5:]} {c['sched_time']}</div>",
                           unsafe_allow_html=True)
            cc[2].markdown(f"<div style='font-size:12px;font-weight:600;color:#059669;padding-top:4px'>{c['expected_revenue']:,}원</div>",
                           unsafe_allow_html=True)
            if cc[3].button("✕", key=f"dc{c['id']}"):
                del_consult(c["id"]); st.rerun()
    else:
        st.markdown("<div style='text-align:center;padding:20px;color:#CBD5E1;font-size:12px'>등록된 상담 없음</div>",
                    unsafe_allow_html=True)

    with st.expander("＋ 상담 추가"):
        with st.form("cf", clear_on_submit=True):
            r1 = st.columns(2)
            cname = r1[0].text_input("이름 *")
            cdate = r1[1].date_input("날짜 *", value=today)
            r2 = st.columns(2)
            ctime = r2[0].text_input("시간", placeholder="14:00")
            crev  = r2[1].number_input("예정매출 (원)", min_value=0, step=10000)
            cnote = st.text_input("메모", placeholder="선택사항")
            if st.form_submit_button("저장 ✓", type="primary", use_container_width=True) and cname.strip():
                add_consult(cname.strip(), cdate.isoformat(), ctime, int(crev), cnote)
                st.rerun()
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# ── 달력 ──────────────────────────────────────────────────────
st.markdown("""<div style='background:white;border-radius:12px;border:1px solid #E2E8F0;
padding:20px 20px 16px'><div style='font-size:14px;font-weight:700;color:#0F172A;
margin-bottom:14px'>🗓️ 이번 달 상담 일정</div>""", unsafe_allow_html=True)

yr, mo = today.year, today.month
cmap = {}
for c in consults:
    if c["sched_date"][:7] == f"{yr:04d}-{mo:02d}":
        d = int(c["sched_date"][8:])
        cmap.setdefault(d, []).append(c["name"])

day_hdr = ["월","화","수","목","금","토","일"]
html = "<table style='width:100%;border-collapse:collapse;table-layout:fixed'><tr>"
for i, dl in enumerate(day_hdr):
    clr = "#EF4444" if i==6 else ("#3B82F6" if i==5 else "#64748B")
    html += f"<th style='padding:6px 2px;font-size:11px;font-weight:700;color:{clr};text-align:center;border-bottom:2px solid #F1F5F9'>{dl}</th>"
html += "</tr>"

for week in cal_lib.monthcalendar(yr, mo):
    html += "<tr>"
    for wi, day in enumerate(week):
        if day == 0:
            html += "<td style='height:56px;padding:3px'></td>"; continue
        is_td  = day == today.day
        is_sun = wi == 6
        is_sat = wi == 5
        num_c  = "white" if is_td else ("#EF4444" if is_sun else ("#3B82F6" if is_sat else "#374151"))
        bg     = "#4F46E5" if is_td else "transparent"
        events = cmap.get(day, [])
        dots = "".join(f"<span style='display:inline-block;width:5px;height:5px;border-radius:50%;background:{'rgba(255,255,255,.85)' if is_td else '#4F46E5'};margin-right:2px'></span>" for _ in events[:3])
        names = "".join(f"<div style='font-size:9px;color:{'rgba(255,255,255,.9)' if is_td else '#4F46E5'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%'>{n}</div>" for n in events[:2])
        html += f"""<td style='padding:3px;vertical-align:top;border:1px solid #F8FAFC'>
<div style='background:{bg};border-radius:7px;padding:4px 5px;min-height:52px'>
  <div style='font-size:12px;font-weight:{"800" if is_td else "500"};color:{num_c};margin-bottom:2px'>{day}</div>
  {dots}{names}
</div></td>"""
    html += "</tr>"
html += "</table>"
st.markdown(html, unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# ── 시간표 배정 자동화 ─────────────────────────────────────────
st.markdown("""<div style='background:white;border-radius:12px;border:1px solid #E2E8F0;
padding:20px 20px 8px'><div style='font-size:14px;font-weight:700;color:#0F172A'>📋 시간표 배정 자동화</div>
<div style='font-size:12px;color:#94A3B8;margin-top:3px;margin-bottom:16px'>
정보 입력 → 배정 문구 자동 생성 후 복사</div></div>""", unsafe_allow_html=True)

gtype = st.selectbox("배정 유형 선택", ["신규 배정","과목변경 배정","배정 취소","날짜변경 배정"],
                     label_visibility="collapsed")

result = ""

if gtype == "신규 배정":
    with st.form("g1", clear_on_submit=False):
        c1,c2,c3,c4 = st.columns(4)
        nd   = c1.date_input("날짜 *", value=today)
        nt   = c2.text_input("시간 *", placeholder="12:00")
        ns   = c3.text_input("과목명 *", placeholder="스케치업2")
        nts  = c4.text_input("시간대", placeholder="주말  /  평일")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("신규", nd=nd.isoformat(), nt=nt, ns=ns, nts=nts)

elif gtype == "과목변경 배정":
    with st.form("g2", clear_on_submit=False):
        st.markdown("<div style='font-size:11px;color:#94A3B8;margin-bottom:4px'>▸ 이전</div>", unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        os_  = c1.text_input("이전 과목명", placeholder="캐드2")
        ofee = c2.number_input("이전 수강료 (0=없음)", min_value=0, step=10000)
        st.markdown("<div style='font-size:11px;color:#94A3B8;margin:8px 0 4px'>▸ 변경 후</div>", unsafe_allow_html=True)
        c3,c4,c5,c6 = st.columns(4)
        nd   = c3.date_input("날짜 *", value=today)
        nt   = c4.text_input("시간 *", placeholder="19:00")
        ns   = c5.text_input("과목명 *", placeholder="실내건축이론1")
        nts  = c6.text_input("시간대", placeholder="주말")
        c7,c8 = st.columns(2)
        nfee = c8.number_input("새 수강료 (0=없음)", min_value=0, step=10000)
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("과목변경", os_=os_, ofee=int(ofee),
                              nd=nd.isoformat(), nt=nt, ns=ns, nts=nts, nfee=int(nfee))

elif gtype == "배정 취소":
    with st.form("g3", clear_on_submit=False):
        c1,c2,c3 = st.columns(3)
        od  = c1.date_input("날짜 *", value=today)
        ot  = c2.text_input("시간 *", placeholder="19:00")
        os_ = c3.text_input("과목명 *", placeholder="실내건축이론1")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("취소", od=od.isoformat(), ot=ot, os_=os_)

elif gtype == "날짜변경 배정":
    with st.form("g4", clear_on_submit=False):
        st.markdown("<div style='font-size:11px;color:#94A3B8;margin-bottom:4px'>▸ 이전 일정</div>", unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        od  = c1.date_input("날짜", value=today, key="od")
        ot  = c2.text_input("시간", placeholder="19:00", key="ot")
        os_ = c3.text_input("과목명", placeholder="캐드2")
        st.markdown("<div style='font-size:11px;color:#94A3B8;margin:8px 0 4px'>▸ 새 일정</div>", unsafe_allow_html=True)
        c4,c5 = st.columns([1,1])
        nd = c4.date_input("날짜", value=today, key="nd")
        nt = c5.text_input("시간", placeholder="19:00", key="nt")
        if st.form_submit_button("✨ 문구 생성", use_container_width=True):
            result = gen_text("날짜변경", od=od.isoformat(), ot=ot, os_=os_,
                              nd=nd.isoformat(), nt=nt)

if result:
    ss.assign_out = result

if ss.assign_out:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.text_area("📋 생성된 배정 문구 (전체 선택 후 복사)", value=ss.assign_out, height=90)
    if st.button("🗑 초기화"):
        ss.assign_out = ""; st.rerun()

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
