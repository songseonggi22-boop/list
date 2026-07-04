import streamlit as st
import sqlite3, time, calendar as cal_lib
from datetime import date, datetime, timedelta

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

def q(sql, a=()):   return get_db().execute(sql, a).fetchall()
def run(sql, a=()):  get_db().execute(sql, a); get_db().commit()

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
    run("""INSERT INTO daily_log(log_date, team_name, rep1_name, rep1_pct, rep2_name, ddaz_den, month_target)
           VALUES(?,?,?,?,?,?,?)""",
        (d, base.get("team_name", "2-3팀"), base.get("rep1_name", ""), base.get("rep1_pct", 60),
         base.get("rep2_name", ""), base.get("ddaz_den", 32), base.get("month_target", 0)))
    return get_log(d)

def save_log(d, **vals):
    sets = ",".join(f"{k}=?" for k in vals)
    run(f"UPDATE daily_log SET {sets} WHERE log_date=?", (*vals.values(), d))

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

    assignees = sorted({t["assignee"] for t in tasks if t["assignee"]})
    view = st.selectbox("보기", ["전체"] + assignees, label_visibility="collapsed")
    view_tasks = tasks if view == "전체" else [t for t in tasks if t["assignee"] == view]

    CATS = [
        ("우선순위1","우선순위 1","kp1"),
        ("우선순위2","우선순위 2","kp2"),
        ("주중업무", "주중업무",  "kp3"),
    ]
    k1, k2, k3 = st.columns(3, gap="small")

    for col_w, (cat_key, cat_label, kp_cls) in zip([k1, k2, k3], CATS):
        with col_w:
            st.markdown(f'<div class="k-head {kp_cls}">{cat_label}</div>', unsafe_allow_html=True)

            cat_tasks = [t for t in view_tasks if t["category"] == cat_key]
            if not cat_tasks:
                st.markdown('<div style="color:#ccc;font-size:11px;text-align:center;padding:12px 0">할 일 없음</div>',
                            unsafe_allow_html=True)

            for t in cat_tasks:
                is_done   = bool(t["is_done"])
                d_display = t["task_date"][5:].replace("-",".") if t["task_date"] else ""
                tc        = "done" if is_done else ""
                a_display = t["assignee"]

                st.markdown(f"""
<div class="t-card">
  {'<div class="t-date">'+d_display+(' · '+a_display if a_display else '')+'</div>' if (d_display or a_display) else ''}
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
                r = st.columns(2)
                nd_val = r[0].date_input("날짜", value=today,
                                       label_visibility="collapsed", key=f"nd_{cat_key}")
                new_assignee = r[1].text_input("담당자", placeholder="담당자",
                                       label_visibility="collapsed", key=f"as_{cat_key}")
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

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── 영업일지 자동 생성 (달력 상담/매출 자동계산 + 수동 항목) ──────
st.markdown("""
<div class="db-card">
  <div class="db-card-title">📈 영업일지 자동 생성</div>
  <div style='font-size:12px;color:#999;margin-top:-10px;margin-bottom:14px'>
    상담건수·매출은 캘린더에서 자동 계산됩니다. 나머지 수치만 입력하세요.
  </div>
</div>""", unsafe_allow_html=True)

with st.expander("📝 일지 정보 입력 (수동 항목)"):
    with st.form("logf"):
        c1, c2, c3, c4 = st.columns(4)
        team_name = c1.text_input("팀명", value=log["team_name"])
        rep1_name = c2.text_input("담당자1(매출귀속)", value=log["rep1_name"])
        rep1_pct  = c3.number_input("귀속%", value=log["rep1_pct"], step=5)
        rep2_name = c4.text_input("담당자2", value=log["rep2_name"])

        c5, c6, c7, c8 = st.columns(4)
        rep1_call = c5.text_input("담당자1 통화시간", value=log["rep1_call"], placeholder="00:51:05")
        rep2_call = c6.text_input("담당자2 통화시간", value=log["rep2_call"], placeholder="00:38:14")
        done_count = c7.number_input("15시 기준 완료건", value=log["done_count"], min_value=0)
        interview_count = c8.number_input("면접예정(건)", value=log["interview_count"], min_value=0)

        c9, c10, c11, c12 = st.columns(4)
        registered   = c9.number_input("등록", value=log["registered"], min_value=0)
        cod          = c10.number_input("COD", value=log["cod"], min_value=0)
        unregistered = c11.number_input("미등록", value=log["unregistered"], min_value=0)
        refund       = c12.number_input("환불(원)", value=log["refund"], min_value=0, step=10000)

        c13, c14, c15, c16 = st.columns(4)
        actual_revenue = c13.number_input("금일매출결과(원)", value=log["actual_revenue"], min_value=0, step=10000)
        tmr_target     = c14.number_input("익일목표매출(원)", value=log["tmr_target"], min_value=0, step=10000)
        ddaz_num       = c15.number_input("따즈아 (분자)", value=log["ddaz_num"], min_value=0)
        ddaz_den       = c16.number_input("따즈아 (분모)", value=log["ddaz_den"], min_value=0)

        c17, c18 = st.columns(2)
        month_target   = c17.number_input(f"{today.month}월 팀목표매출(원)", value=log["month_target"], min_value=0, step=100000)
        month_achieved = c18.number_input("현재달성매출(원)", value=log["month_achieved"], min_value=0, step=100000)

        if st.form_submit_button("저장", use_container_width=True):
            save_log(today_str, team_name=team_name, rep1_name=rep1_name, rep1_pct=int(rep1_pct),
                      rep2_name=rep2_name, rep1_call=rep1_call, rep2_call=rep2_call,
                      done_count=int(done_count), interview_count=int(interview_count),
                      registered=int(registered), cod=int(cod), unregistered=int(unregistered),
                      refund=int(refund), actual_revenue=int(actual_revenue), tmr_target=int(tmr_target),
                      ddaz_num=int(ddaz_num), ddaz_den=int(ddaz_den),
                      month_target=int(month_target), month_achieved=int(month_achieved))
            st.rerun()

t1, t2, t3 = st.tabs(["출근보고", "15시보고", "마감보고"])

with t1:
    st.code(f"""{log['team_name']} 영업일지({today.month:02d}.{today.day:02d})

- 금일 입금예정: {today_rev // 10000}만원

{today_rev // 10000}만원 / {log['rep1_name']} {log['rep1_pct']}%

- 금일 상담건수 : {today_cnt}건(정규{today_reg}건/단과{today_dan}건)
- 면접예정: {log['interview_count']}건""", language=None)

with t2:
    st.code(f"""[{log['team_name']} 15:00 보고]
{log['done_count']} / {today_cnt}
익일상담 {tmr_cnt} / 익일예정 {tmr_rev // 10000}
모레상담 {daf_cnt} / 모레예정 {daf_rev // 10000}
익일면접 {log['interview_count']}건
따즈아 {log['ddaz_num']} / {log['ddaz_den']}""", language=None)

with t3:
    pct = round(log['month_achieved'] / log['month_target'] * 100) if log['month_target'] else 0
    st.code(f"""컴퓨터 {log['team_name']} 영업마감보고

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
현재달성율 : {pct}%""", language=None)
