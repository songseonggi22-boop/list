import streamlit as st
import sqlite3, time, json
from datetime import date, datetime

DB = "teamboard.db"

CATS = ["우선순위1","우선순위2","주중업무","홈관리","LIFE","공부&독서"]
CAT_ACC = {
    "우선순위1": "#DC2626", "우선순위2": "#EA580C",
    "주중업무":  "#374151", "홈관리":   "#1D4ED8",
    "LIFE":     "#7C3AED", "공부&독서": "#B45309",
}
CAT_BG = {
    "우선순위1": "#FEF2F2", "우선순위2": "#FFF7ED",
    "주중업무":  "#F9FAFB", "홈관리":   "#F9FAFB",
    "LIFE":     "#F9FAFB", "공부&독서": "#F9FAFB",
}
CAT_EMO = {
    "우선순위1":"🔴","우선순위2":"🟠","주중업무":"💼",
    "홈관리":"🏠","LIFE":"🌿","공부&독서":"📚",
}

TAG_PAL = [
    ("#EDE9FE","#5B21B6"),("#FEE2E2","#991B1B"),("#FEF3C7","#92400E"),
    ("#D1FAE5","#065F46"),("#DBEAFE","#1E40AF"),("#FCE7F3","#831843"),
    ("#FEF9C3","#713F12"),("#E0F2FE","#0369A1"),
]
def tag_color(tag):
    return TAG_PAL[sum(ord(c) for c in tag) % len(TAG_PAL)]

# ── DB ────────────────────────────────────────────
@st.cache_resource
def get_db():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS tasks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        category TEXT DEFAULT '주중업무',
        status TEXT DEFAULT 'todo',
        due_date TEXT DEFAULT '',
        tags TEXT DEFAULT '[]',
        created_at INTEGER DEFAULT 0
    )""")
    # migrate old schema
    cols = [r[1] for r in c.execute("PRAGMA table_info(tasks)").fetchall()]
    if "category" not in cols:
        c.execute("ALTER TABLE tasks ADD COLUMN category TEXT DEFAULT '주중업무'")
        # map old priority to category
        if "priority" in cols:
            c.execute("UPDATE tasks SET category='우선순위1' WHERE priority='urgent'")
            c.execute("UPDATE tasks SET category='우선순위2' WHERE priority='high'")
    c.commit()
    return c

def qtasks():
    rows = get_db().execute(
        "SELECT id,title,category,status,due_date,tags,created_at FROM tasks ORDER BY created_at DESC"
    ).fetchall()
    keys = "id title category status due_date tags created_at".split()
    return [dict(zip(keys, r)) for r in rows]

def add_task(title, category, due_date, tags):
    get_db().execute(
        "INSERT INTO tasks(title,category,status,due_date,tags,created_at) VALUES(?,?,?,?,?,?)",
        (title, category, "todo", due_date, json.dumps(tags, ensure_ascii=False), int(time.time()*1000))
    )
    get_db().commit()

def set_done(tid, done):
    get_db().execute("UPDATE tasks SET status=? WHERE id=?", ("done" if done else "todo", tid))
    get_db().commit()

def del_task(tid):
    get_db().execute("DELETE FROM tasks WHERE id=?", (tid,))
    get_db().commit()

def fmt_due(d):
    if not d: return ""
    td = date.today().isoformat()
    if d == td: return "오늘 ⚡"
    if d < td:  return d[5:].replace("-","/") + " ⚠"
    return d[5:].replace("-","/")

# ── Page ──────────────────────────────────────────
st.set_page_config(page_title="To Do List", page_icon="✅", layout="wide",
                   initial_sidebar_state="collapsed")

ss = st.session_state
if "add_cat" not in ss: ss.add_cat = None

st.markdown("""<style>
#MainMenu,footer,[data-testid="stHeader"],[data-testid="stToolbar"]{display:none!important}
.block-container{padding:1.2rem 1.6rem 2rem!important;max-width:100%!important}
[data-testid="stAppViewContainer"]{background:#F7F6F3}
.main{background:#F7F6F3}
[data-testid="stButton"]>button{
  padding:4px 10px!important;font-size:12px!important;border-radius:5px!important;
  font-weight:500!important;min-height:0!important;line-height:1.4!important;
  border:1px solid #E5E7EB!important;background:white!important;color:#6B7280!important}
[data-testid="stButton"]>button:hover{background:#F3F4F6!important;border-color:#D1D5DB!important}
[data-testid="stForm"]{border:none!important;padding:0!important}
.stTextInput>div>div>input{border-radius:6px!important;font-size:12px!important;padding:4px 8px!important}
.stSelectbox>div>div{border-radius:6px!important;font-size:12px!important}
.stDateInput>div>div>input{border-radius:6px!important;font-size:12px!important}
[data-testid="stCheckbox"] label span{font-size:12px!important}
div[data-testid="stHorizontalBlock"]{gap:0.6rem!important}
</style>""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────
tasks = qtasks()
today = date.today()
total  = len(tasks)
done_n = sum(1 for t in tasks if t["status"]=="done")
left_n = total - done_n
now    = datetime.now().strftime("%H:%M")
weekday = ["월","화","수","목","금","토","일"][today.weekday()]

h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(f"""
<div style='padding:4px 0 12px'>
  <div style='font-size:24px;font-weight:700;color:#111827;letter-spacing:-0.5px'>TO DO LIST</div>
  <div style='font-size:13px;color:#9CA3AF;margin-top:3px'>
    {today.strftime('%Y. %m. %d')} ({weekday}) &nbsp;·&nbsp;
    전체 <b style='color:#374151'>{total}</b> &nbsp;·&nbsp;
    완료 <b style='color:#059669'>{done_n}</b> &nbsp;·&nbsp;
    남은 할 일 <b style='color:#DC2626'>{left_n}</b>
  </div>
</div>""", unsafe_allow_html=True)
with h2:
    st.markdown(f"""
<div style='text-align:right;padding-top:4px'>
  <div style='font-size:40px;font-weight:200;color:#1A1A1A;letter-spacing:6px;font-family:monospace'>{now}</div>
</div>""", unsafe_allow_html=True)

# ── Kanban ────────────────────────────────────────
cols = st.columns(6, gap="small")

for ci, cat in enumerate(CATS):
    with cols[ci]:
        cat_tasks = [t for t in tasks if t["category"] == cat]
        done_c = sum(1 for t in cat_tasks if t["status"] == "done")
        acc = CAT_ACC[cat]
        bg  = CAT_BG[cat]
        emo = CAT_EMO[cat]

        # Column header
        st.markdown(f"""
<div style='background:{bg};border:1px solid #E5E7EB;border-top:3px solid {acc};
            border-radius:8px 8px 0 0;padding:10px 12px 8px'>
  <div style='font-size:12px;font-weight:700;color:{acc}'>{emo} {cat}</div>
  <div style='font-size:10px;color:#9CA3AF;margin-top:2px'>{len(cat_tasks)}개 · 완료 {done_c}</div>
</div>""", unsafe_allow_html=True)

        # Add button
        if st.button("＋ 추가", key=f"btn_{cat}", use_container_width=True):
            ss.add_cat = cat if ss.add_cat != cat else None
            st.rerun()

        # Add form
        if ss.add_cat == cat:
            with st.form(f"f_{cat}", clear_on_submit=True):
                title   = st.text_input("제목", placeholder="할 일 입력...")
                due     = st.date_input("마감일", value=None, label_visibility="collapsed")
                tags_in = st.text_input("태그", placeholder="태그1, 태그2")
                ok, cnl = st.columns(2)
                if ok.form_submit_button("추가 ✓", use_container_width=True) and title.strip():
                    tag_list = [x.strip() for x in tags_in.split(",") if x.strip()]
                    add_task(title.strip(), cat, due.isoformat() if due else "", tag_list)
                    ss.add_cat = None
                    st.rerun()
                if cnl.form_submit_button("취소", use_container_width=True):
                    ss.add_cat = None
                    st.rerun()

        # Task cards
        if not cat_tasks:
            st.markdown("""<div style='text-align:center;padding:24px 0;color:#D1D5DB;font-size:11px'>
              할 일 없음
            </div>""", unsafe_allow_html=True)
        else:
            for t in cat_tasks:
                is_done  = t["status"] == "done"
                dl       = fmt_due(t["due_date"])
                is_over  = t["due_date"] and t["due_date"] < today.isoformat() and not is_done

                try:    tag_list = json.loads(t["tags"]) if t["tags"] else []
                except: tag_list = []

                tags_html = "".join(
                    f"<span style='background:{tag_color(g)[0]};color:{tag_color(g)[1]};"
                    f"border-radius:4px;padding:1px 6px;font-size:10px;font-weight:600;"
                    f"margin-right:3px;white-space:nowrap'>{g}</span>"
                    for g in tag_list[:3]
                )
                due_color  = "#DC2626" if is_over else "#9CA3AF"
                title_sty  = "text-decoration:line-through;color:#D1D5DB" if is_done else "color:#111827;font-weight:500"
                card_border = "1px solid #E5E7EB"

                st.markdown(f"""
<div style='background:white;border:{card_border};border-radius:6px;
            padding:10px 12px 8px;margin:4px 0'>
  <div style='font-size:12px;{title_sty};line-height:1.5'>{t['title']}</div>
  {f"<div style='font-size:10px;color:{due_color};margin-top:3px'>{dl}</div>" if dl else ""}
  {f"<div style='margin-top:5px'>{tags_html}</div>" if tags_html else ""}
</div>""", unsafe_allow_html=True)

                c1, c2 = st.columns([1, 0.4])
                chk = c1.checkbox("완료", value=is_done, key=f"chk_{t['id']}")
                if chk != is_done:
                    set_done(t["id"], chk)
                    st.rerun()
                if c2.button("✕", key=f"del_{t['id']}"):
                    del_task(t["id"])
                    st.rerun()
