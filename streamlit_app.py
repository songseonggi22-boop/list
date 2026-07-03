import streamlit as st
import sqlite3, json, time, socket
from datetime import date

# ── 상수 ──────────────────────────────────────────
DB = "teamboard.db"
PRIORITY = {
    "urgent": ("🔴","긴급","#FEE2E2","#991B1B"),
    "high":   ("🟠","높음","#FFF7ED","#9A3412"),
    "normal": ("🟣","보통","#EDE9FE","#5B21B6"),
    "low":    ("⚫","낮음","#F1F5F9","#64748B"),
}
STATUS = {
    "todo":  ("⭕","할 일", "#EDE9FE","#5B21B6"),
    "doing": ("🔄","진행 중","#FEF3C7","#92400E"),
    "done":  ("✅","완료",  "#D1FAE5","#065F46"),
}
NEXT_ST = {"todo":"doing","doing":"done","done":"todo"}
AVC = ["#5B4CF5","#EC4899","#059669","#0EA5E9","#D97706","#DC2626","#7C3AED","#0891B2"]
def ncolor(n): return AVC[sum(ord(c) for c in n) % len(AVC)] if n else "#9BA3B5"

# ── DB ────────────────────────────────────────────
@st.cache_resource
def db():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS tasks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, assignee TEXT DEFAULT '',
        status TEXT DEFAULT 'todo', priority TEXT DEFAULT 'normal',
        due_date TEXT DEFAULT '', note TEXT DEFAULT '',
        tags TEXT DEFAULT '[]', created_at INTEGER DEFAULT 0, sort_order INTEGER DEFAULT 0
    )"""); c.commit(); return c

def qtasks():
    r = db().execute("SELECT id,title,assignee,status,priority,due_date,note,tags,created_at,sort_order FROM tasks ORDER BY sort_order,created_at DESC").fetchall()
    k = "id title assignee status priority due_date note tags created_at sort_order".split()
    return [dict(zip(k,row)) for row in r]

def ins(title,assignee,status,priority,due_date,note):
    ts = int(time.time()*1000)
    mx = db().execute("SELECT COALESCE(MAX(sort_order),0) FROM tasks").fetchone()[0]
    db().execute("INSERT INTO tasks(title,assignee,status,priority,due_date,note,created_at,sort_order) VALUES(?,?,?,?,?,?,?,?)",
                 (title,assignee,status,priority,due_date,note,ts,mx+1)); db().commit()

def upd(tid,**kw):
    db().execute("UPDATE tasks SET "+", ".join(f"{k}=?" for k in kw)+" WHERE id=?", list(kw.values())+[tid]); db().commit()

def rem(tid):
    db().execute("DELETE FROM tasks WHERE id=?", (tid,)); db().commit()

def local_ip():
    try:
        s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM); s.connect(("8.8.8.8",80)); ip=s.getsockname()[0]; s.close(); return ip
    except: return "localhost"

def fmt_due(d):
    if not d: return ""
    td = date.today().isoformat()
    if d == td: return "오늘 ⚡"
    if d < td:  return d[5:].replace("-","/")+" ⚠"
    return d[5:].replace("-","/")

# ── 페이지 설정 ───────────────────────────────────
st.set_page_config(page_title="Team Board", page_icon="📋", layout="wide", initial_sidebar_state="collapsed")
ss = st.session_state
for k,v in [("sort","newest"),("fpri","all"),("q",""),("form",False),("eid",None),("prea",""),("clp",{})]:
    if k not in ss: ss[k] = v

# ── CSS ───────────────────────────────────────────
st.markdown("""<style>
#MainMenu,footer,[data-testid="stHeader"],[data-testid="stToolbar"]{display:none!important}
.block-container{padding:.5rem 1.5rem!important;max-width:100%!important}
[data-testid="stButton"]>button{
  padding:5px 14px!important;font-size:12px!important;border-radius:20px!important;
  font-weight:600!important;min-height:0!important;height:auto!important;line-height:1.4!important}
[data-testid="stButton"]>button[kind="primary"]{
  background:linear-gradient(135deg,#5B4CF5,#EC4899)!important;border:none!important;color:white!important}
[data-testid="metric-container"]{background:white;border-radius:10px;padding:12px 14px;border:1px solid #DDE1EA}
hr{margin:4px 0!important;border-color:#EAECF3!important;opacity:.5!important}
[data-testid="stCheckbox"]>label>div:first-child{width:22px!important;height:22px!important;border-radius:6px!important}
.stTextInput>div>div>input,.stSelectbox>div>div{border-radius:10px!important}
[data-testid="stExpander"]{border:1.5px solid #DDE1EA!important;border-radius:12px!important;background:white!important}
[data-testid="stForm"]{border:none!important}
</style>""", unsafe_allow_html=True)

# ── 데이터 로드 & 통계 ────────────────────────────
tasks = qtasks()
vt = [t for t in tasks
      if (ss.fpri=="all" or t["priority"]==ss.fpri)
      and (not ss.q or ss.q.lower() in (t["title"]+t["assignee"]).lower())]

total  = len(tasks)
done_c = sum(1 for t in tasks if t["status"]=="done")
doing_c= sum(1 for t in tasks if t["status"]=="doing")
todo_c = sum(1 for t in tasks if t["status"]=="todo")
over_c = sum(1 for t in tasks if t["due_date"] and t["due_date"]<date.today().isoformat() and t["status"]!="done")
pct    = round(done_c/total*100) if total else 0

# ── 짱구 (할 일 남아있을 때 달림) ─────────────────
if total > 0 and done_c < total:
    st.components.v1.html("""
<div style="width:100%;height:72px;overflow:hidden;position:relative">
<div style="position:absolute;top:6px;left:0;width:56px;animation:run 10s linear infinite;filter:drop-shadow(1px 3px 4px rgba(0,0,0,.22))">
<div style="display:flex;flex-direction:column;align-items:center;animation:bob .25s ease-in-out infinite alternate">
  <div style="width:40px;height:19px;background:#111;border-radius:22px 22px 0 0;margin-bottom:-4px;position:relative;z-index:3"></div>
  <div style="width:38px;height:34px;background:#f9c47a;border-radius:50% 50% 42% 42%/56% 56% 44% 44%;position:relative">
    <div style="position:absolute;top:9px;left:4px;width:10px;height:3px;background:#222;border-radius:2px;transform:rotate(12deg)"></div>
    <div style="position:absolute;top:9px;right:4px;width:10px;height:3px;background:#222;border-radius:2px;transform:rotate(-12deg)"></div>
    <div style="position:absolute;top:15px;left:7px;width:6px;height:8px;background:#111;border-radius:50%"></div>
    <div style="position:absolute;top:15px;right:7px;width:6px;height:8px;background:#111;border-radius:50%"></div>
    <div style="position:absolute;bottom:3px;left:50%;transform:translateX(-50%);width:24px;height:10px;background:#bf3020;border-radius:2px 2px 12px 12px;overflow:hidden">
      <div style="width:100%;height:5px;background:#fff"></div></div>
  </div>
  <div style="width:30px;height:15px;background:#fff;border:1.5px solid #ddd;border-bottom:none;border-radius:3px 3px 0 0;margin:0 auto"></div>
  <div style="width:32px;height:12px;background:#f5c518;margin:0 auto;border-radius:0 0 3px 3px"></div>
  <div style="display:flex;justify-content:center;gap:3px;margin-top:2px">
    <div style="width:11px;height:13px;background:#f9c47a;border-radius:3px;transform-origin:top center;animation:lL .25s ease-in-out infinite alternate"></div>
    <div style="width:11px;height:13px;background:#f9c47a;border-radius:3px;transform-origin:top center;animation:lR .25s ease-in-out infinite alternate"></div>
  </div>
</div></div>
<style>
@keyframes run{from{transform:translateX(-80px)}to{transform:translateX(calc(100vw + 80px))}}
@keyframes bob{from{transform:translateY(0) rotate(-3deg)}to{transform:translateY(-5px) rotate(3deg)}}
@keyframes lL{from{transform:rotate(-32deg)}to{transform:rotate(28deg)}}
@keyframes lR{from{transform:rotate(28deg)}to{transform:rotate(-32deg)}}
</style></div>""", height=72)

# ── 헤더 ──────────────────────────────────────────
hc = st.columns([4, 2.5, 1.5])
hc[0].markdown("### 📋 Team Board")
ss.q = hc[1].text_input("search", placeholder="🔍 할 일 / 담당자 검색...",
                         label_visibility="collapsed", value=ss.q)
if hc[2].button("＋ 새 할 일", type="primary", use_container_width=True):
    ss.form=True; ss.eid=None; ss.prea=""; st.rerun()

# ── 통계 ──────────────────────────────────────────
mc = st.columns(6)
mc[0].metric("전체",   total)
mc[1].metric("대기",   todo_c)
mc[2].metric("진행",   doing_c)
mc[3].metric("완료",   done_c)
mc[4].metric("기한초과", over_c)
with mc[5]:
    st.write("**팀 진행률**"); st.progress(pct/100); st.caption(f"{pct}%")
st.divider()

# ── 툴바 (필터 + 정렬) ────────────────────────────
tc = st.columns([5, 4, 0.8])
with tc[0]:
    st.caption("**우선순위**")
    pc = st.columns(5)
    for i,(k,v) in enumerate([("all","전체"),("urgent","🔴 긴급"),("high","🟠 높음"),("normal","🟣 보통"),("low","⚫ 낮음")]):
        if pc[i].button(v, key=f"fp{k}", use_container_width=True,
                        type="primary" if ss.fpri==k else "secondary"):
            ss.fpri=k; st.rerun()
with tc[1]:
    st.caption("**정렬**")
    sc = st.columns(3)
    for i,(k,v) in enumerate([("newest","최신순"),("alpha","가나다순"),("manual","직접정렬")]):
        if sc[i].button(v, key=f"sm{k}", use_container_width=True,
                        type="primary" if ss.sort==k else "secondary"):
            ss.sort=k; st.rerun()
with tc[2]:
    if st.button("🔄 새로고침", use_container_width=True):
        st.rerun()

# ── 추가/수정 폼 ───────────────────────────────────
if ss.form:
    et = next((t for t in tasks if t["id"]==ss.eid), None) if ss.eid else None
    with st.expander("📝 " + ("할 일 수정" if et else "새 할 일 추가"), expanded=True):
        with st.form("tf", clear_on_submit=True):
            title = st.text_input("제목 *", value=et["title"] if et else "")
            f1, f2 = st.columns(2)
            assignee = f1.text_input("담당자", value=et["assignee"] if et else ss.prea)
            st_val = f2.selectbox("상태", ["todo","doing","done"],
                index=["todo","doing","done"].index(et["status"] if et else "todo"),
                format_func=lambda x: STATUS[x][1])
            f3, f4 = st.columns(2)
            pri_val = f3.selectbox("우선순위", ["urgent","high","normal","low"],
                index=["urgent","high","normal","low"].index(et["priority"] if et else "normal"),
                format_func=lambda x: PRIORITY[x][0]+" "+PRIORITY[x][1])
            try: dd_def = date.fromisoformat(et["due_date"]) if et and et["due_date"] else None
            except: dd_def = None
            due = f4.date_input("마감일", value=dd_def)
            note = st.text_area("메모", value=et["note"] if et else "", height=70)
            bs, bc_ = st.columns(2)
            ok  = bs.form_submit_button("저장 ✓", type="primary", use_container_width=True)
            cnl = bc_.form_submit_button("취소", use_container_width=True)
            if ok and title.strip():
                ds = due.isoformat() if due else ""
                if et: upd(et["id"], title=title, assignee=assignee, status=st_val, priority=pri_val, due_date=ds, note=note)
                else:  ins(title, assignee, st_val, pri_val, ds, note)
                ss.form=False; ss.eid=None; ss.prea=""; st.rerun()
            if cnl:
                ss.form=False; ss.eid=None; ss.prea=""; st.rerun()

# ── 섹션 정렬 함수 ────────────────────────────────
def sort_sec(lst):
    if ss.sort=="alpha":  return sorted(lst, key=lambda t: t["title"])
    if ss.sort=="newest": return sorted(lst, key=lambda t: t["created_at"], reverse=True)
    return sorted(lst, key=lambda t: t["sort_order"])

# ── 담당자 섹션 ───────────────────────────────────
names = list(dict.fromkeys(t["assignee"] for t in vt if t["assignee"]))
if any(not t["assignee"] for t in vt): names.append("")
today = date.today().isoformat()

for name in names:
    sec = [t for t in vt if t["assignee"]==name]
    if not sec: continue

    ck   = f"c_{name}"
    open_ = not ss.clp.get(ck, False)
    done_n  = sum(1 for t in sec if t["status"]=="done")
    doing_n = sum(1 for t in sec if t["status"]=="doing")
    todo_n  = sum(1 for t in sec if t["status"]=="todo")
    sp   = round(done_n/len(sec)*100) if sec else 0
    c    = ncolor(name)
    ini  = (name or "?")[0].upper()
    chips = ""
    if todo_n:  chips += f"<span style='background:#EDE9FE;color:#5B21B6;border-radius:4px;padding:1px 7px;font-size:10px;font-weight:700;margin-right:4px'>대기 {todo_n}</span>"
    if doing_n: chips += f"<span style='background:#FEF3C7;color:#92400E;border-radius:4px;padding:1px 7px;font-size:10px;font-weight:700;margin-right:4px'>진행 {doing_n}</span>"
    if done_n:  chips += f"<span style='background:#D1FAE5;color:#065F46;border-radius:4px;padding:1px 7px;font-size:10px;font-weight:700'>완료 {done_n}</span>"

    # 섹션 헤더
    sh = st.columns([0.18, 5, 1.8])
    with sh[0]:
        if st.button("▼" if open_ else "▶", key=f"tog{name}"):
            ss.clp[ck] = not ss.clp.get(ck, False); st.rerun()
    with sh[1]:
        st.markdown(f"""<div style="display:flex;align-items:center;gap:10px;padding:6px 0 4px;border-left:4px solid {c};padding-left:12px">
          <div style="width:34px;height:34px;border-radius:50%;background:{c};display:flex;align-items:center;justify-content:center;color:white;font-weight:800;font-size:15px;flex-shrink:0">{ini}</div>
          <div>
            <span style="font-size:15px;font-weight:700;color:#1A1B2E">{name or '미지정'}</span>
            <span style="font-size:11px;color:#9BA3B5;margin-left:8px">{done_n}/{len(sec)} 완료 ({sp}%)</span>
            <div style="margin-top:5px">{chips}</div>
          </div>
        </div>""", unsafe_allow_html=True)
        st.progress(sp/100)
    with sh[2]:
        if name and st.button(f"＋ {name}에게 추가", key=f"add{name}"):
            ss.form=True; ss.eid=None; ss.prea=name; st.rerun()

    if not open_: continue

    # 컬럼 헤더
    ch = st.columns([0.55, 0.4, 4, 1.4, 1, 1, 0.8])
    for col, lbl in zip(ch, ["","","할 일","상태","우선순위","마감일",""]):
        col.markdown(f"<p style='font-size:10px;font-weight:700;color:#9BA3B5;text-transform:uppercase;letter-spacing:.5px;margin:4px 0 2px'>{lbl}</p>", unsafe_allow_html=True)

    sorted_sec = sort_sec(sec)
    for i, t in enumerate(sorted_sec):
        is_done = t["status"] == "done"
        is_over = t["due_date"] and t["due_date"] < today and not is_done
        dl = fmt_due(t["due_date"])

        rc = st.columns([0.55, 0.4, 4, 1.4, 1, 1, 0.8])

        # 순서 이동 (직접정렬 모드)
        with rc[0]:
            if ss.sort == "manual":
                if st.button("△", key=f"u{t['id']}") and i > 0:
                    a,b = sorted_sec[i], sorted_sec[i-1]
                    upd(a["id"],sort_order=b["sort_order"]); upd(b["id"],sort_order=a["sort_order"]); st.rerun()
                if st.button("▽", key=f"d{t['id']}") and i < len(sorted_sec)-1:
                    a,b = sorted_sec[i], sorted_sec[i+1]
                    upd(a["id"],sort_order=b["sort_order"]); upd(b["id"],sort_order=a["sort_order"]); st.rerun()

        # 체크박스
        with rc[1]:
            chk = st.checkbox("", value=is_done, key=f"c{t['id']}")
            if chk != is_done:
                upd(t["id"], status="done" if chk else "todo"); st.rerun()

        # 제목 + 메모
        with rc[2]:
            ts_style = "text-decoration:line-through;color:#9BA3B5" if is_done else "color:#1A1B2E;font-weight:500"
            note_html = f"<div style='font-size:11px;color:#9BA3B5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:380px'>{t['note'][:60]}</div>" if t["note"] else ""
            st.markdown(f"<div style='padding:5px 0'><div style='font-size:14px;{ts_style}'>{t['title']}</div>{note_html}</div>", unsafe_allow_html=True)

        # 상태 버튼 (클릭 → 다음 상태)
        with rc[3]:
            si = STATUS[t["status"]]
            if st.button(f"{si[0]} {si[1]}", key=f"st{t['id']}", help="클릭으로 상태 변경"):
                upd(t["id"], status=NEXT_ST[t["status"]]); st.rerun()

        # 우선순위 뱃지
        with rc[4]:
            pi = PRIORITY[t["priority"] or "normal"]
            st.markdown(f"<div style='background:{pi[2]};color:{pi[3]};border-radius:20px;padding:4px 10px;font-size:11px;font-weight:700;text-align:center;margin-top:6px'>{pi[0]} {pi[1]}</div>", unsafe_allow_html=True)

        # 마감일
        with rc[5]:
            due_color = "#DC2626" if is_over else ("#5B4CF5" if t["due_date"]==today else "#5C6070")
            st.markdown(f"<div style='color:{due_color};font-size:12px;font-weight:{'700' if is_over else '500'};margin-top:8px'>{dl}</div>", unsafe_allow_html=True)

        # 수정 / 삭제
        with rc[6]:
            ac = st.columns(2)
            if ac[0].button("✏", key=f"e{t['id']}"):
                ss.form=True; ss.eid=t["id"]; st.rerun()
            if ac[1].button("🗑", key=f"x{t['id']}"):
                rem(t["id"]); st.rerun()

        st.divider()

# ── 빈 상태 ───────────────────────────────────────
if not vt:
    st.markdown("""<div style="text-align:center;padding:60px;color:#9BA3B5">
      <div style="font-size:52px;margin-bottom:12px">📋</div>
      <div style="font-size:17px;font-weight:700;color:#5C6070;margin-bottom:6px">할 일이 없어요</div>
      <div style="font-size:13px">'+ 새 할 일'로 팀 업무를 추가해보세요</div>
    </div>""", unsafe_allow_html=True)

# ── 하단: 팀 공유 정보 ────────────────────────────
st.divider()
ip = local_ip()
st.info(f"**팀 공유:** 팀원 브라우저에서 `http://{ip}:8501` 접속 (같은 와이파이/사무실 네트워크)")
