import streamlit as st
import sqlite3, time, calendar as cal_lib
from datetime import date, datetime

DB = "salesdb.db"

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
        is_done INTEGER DEFAULT 0,
        created_at INTEGER DEFAULT 0);

    CREATE TABLE IF NOT EXISTS consultations(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, sched_date TEXT NOT NULL,
        sched_time TEXT DEFAULT '', expected_revenue INTEGER DEFAULT 0,
        created_at INTEGER DEFAULT 0);
    """)
    c.commit()
    return c

def q(sql, a=()):   return get_db().execute(sql, a).fetchall()
def run(sql, a=()):  get_db().execute(sql, a); get_db().commit()

def get_tasks():
    rows = q("SELECT id,title,category,task_date,is_done FROM tasks ORDER BY is_done,created_at DESC")
    return [dict(zip("id title category task_date is_done".split(), r)) for r in rows]

def get_consults():
    rows = q("SELECT id,name,sched_date,sched_time,expected_revenue FROM consultations ORDER BY sched_date,sched_time")
    return [dict(zip("id name sched_date sched_time expected_revenue".split(), r)) for r in rows]

def add_task(title, cat, d):
    run("INSERT INTO tasks(title,category,task_date,is_done,created_at) VALUES(?,?,?,0,?)",
        (title, cat, d, int(time.time()*1000)))

def toggle_task(tid, v): run("UPDATE tasks SET is_done=? WHERE id=?", (int(v), tid))
def del_task(tid):       run("DELETE FROM tasks WHERE id=?", (tid,))

def add_consult(name, d, t, rev):
    run("INSERT INTO consultations(name,sched_date,sched_time,expected_revenue,created_at) VALUES(?,?,?,?,?)",
        (name, d, t, rev, int(time.time()*1000)))
def del_consult(cid): run("DELETE FROM consultations WHERE id=?", (cid,))

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

today     = date.today()
today_str = today.isoformat()
now_str   = datetime.now().strftime("%H:%M")
wd        = ["월","화","수","목","금","토","일"][today.weekday()]
yr, mo    = today.year, today.month

# ── 데이터 로드 ────────────────────────────────────────────────
tasks    = get_tasks()
consults = get_consults()
today_c  = [c for c in consults if c["sched_date"] == today_str]

# 달력용 이벤트 맵
cmap = {}
for c in consults:
    if c["sched_date"][:7] == f"{yr:04d}-{mo:02d}":
        d = int(c["sched_date"][8:])
        cmap.setdefault(d, []).append((c["name"], c["expected_revenue"]))

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

# ── 헤더 ─────────────────────────────────────────────────────
st.markdown(f"""
<div style='display:flex;justify-content:space-between;align-items:center;
            padding-bottom:14px;border-bottom:2px solid #eaeaea;margin-bottom:20px'>
  <h1 style='font-size:22px;font-weight:700;color:#333'>
    <span style='color:#4f46e5'>■</span> 업무 및 일정 관리 대시보드
  </h1>
  <div style='font-size:13px;color:#666'>
    {today.strftime('%Y년 %m월 %d일')} ({wd}요일)
    &nbsp;<span style='font-size:22px;font-weight:200;font-family:monospace;color:#333'>{now_str}</span>
  </div>
</div>""", unsafe_allow_html=True)

# ── 메인 2컬럼 ────────────────────────────────────────────────
left, right = st.columns([1.05, 1.35], gap="medium")

# ── LEFT: 달력 ───────────────────────────────────────────────
with left:
    ci_cls = ["ci1","ci2","ci3"]
    day_clr = [None,None,None,None,None,"#2f80ed","#ff5b5b"]

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
                f'<span class="ci {ci_cls[i%3]}">{e[0]}</span>'
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
            r2 = st.columns(2)
            ctime = r2[0].text_input("시간", placeholder="14:00")
            crev  = r2[1].number_input("예정매출(원)", min_value=0, step=10000)
            if st.form_submit_button("저장", use_container_width=True) and cname.strip():
                add_consult(cname.strip(), cdate.isoformat(), ctime, int(crev)); st.rerun()

    # 오늘 상담
    if today_c:
        st.markdown("<div style='margin-top:10px;font-size:11px;font-weight:700;color:#4f46e5'>🔵 오늘 상담</div>",
                    unsafe_allow_html=True)
        for c in today_c:
            cc1, cc2 = st.columns([1, 0.13])
            cc1.markdown(f"""
<div style='background:#eef2ff;border-left:3px solid #4f46e5;border-radius:6px;
            padding:7px 12px;margin:3px 0;display:flex;justify-content:space-between;align-items:center'>
  <span style='font-size:12px;font-weight:700;color:#3730a3'>{c['name']}</span>
  <span style='font-size:11px;color:#6366f1'>{c['sched_time']}</span>
  <span style='font-size:11px;font-weight:700;color:#4f46e5'>{c['expected_revenue']:,}원</span>
</div>""", unsafe_allow_html=True)
            if cc2.button("✕", key=f"dc{c['id']}"):
                del_consult(c["id"]); st.rerun()

# ── RIGHT: 칸반 TO DO LIST ────────────────────────────────────
with right:
    st.markdown('<div class="db-card"><div class="db-card-title">☑ TO DO LIST</div>', unsafe_allow_html=True)

    CATS = [
        ("우선순위1","우선순위 1","kp1"),
        ("우선순위2","우선순위 2","kp2"),
        ("주중업무", "주중업무",  "kp3"),
    ]
    k1, k2, k3 = st.columns(3, gap="small")

    for col_w, (cat_key, cat_label, kp_cls) in zip([k1, k2, k3], CATS):
        with col_w:
            st.markdown(f'<div class="k-head {kp_cls}">{cat_label}</div>', unsafe_allow_html=True)

            cat_tasks = [t for t in tasks if t["category"] == cat_key]
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

            # 할 일 추가 폼
            with st.form(f"af_{cat_key}", clear_on_submit=True):
                new_title = st.text_input("", placeholder="＋ 새 페이지",
                                          label_visibility="collapsed", key=f"ti_{cat_key}")
                nd_val = st.date_input("날짜", value=today,
                                       label_visibility="collapsed", key=f"nd_{cat_key}")
                if st.form_submit_button("추가", use_container_width=True) and new_title.strip():
                    add_task(new_title.strip(), cat_key, nd_val.isoformat()); st.rerun()

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

gtype = st.selectbox("배정 유형", ["신규 배정","과목변경 배정","배정 취소","날짜변경 배정"],
                     label_visibility="collapsed")
result = ""

if gtype == "신규 배정":
    with st.form("g1"):
        c1,c2,c3,c4 = st.columns(4)
        nd  = c1.date_input("날짜*", value=today)
        nt  = c2.text_input("시간*", placeholder="12:00")
        ns  = c3.text_input("과목명*", placeholder="스케치업2")
        nts = c4.text_input("시간대", placeholder="주말")
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
    st.text_area("📋 생성된 배정 문구 (복사하세요)", value=ss.assign_out, height=80)
    if st.button("🗑 초기화"):
        ss.assign_out = ""; st.rerun()
