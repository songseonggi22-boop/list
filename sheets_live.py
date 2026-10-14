"""링크 공개 구글시트 실시간 읽기 — 격일 달력시간표 · 비대면 디스코드 초대코드.
인증 없이 export?format=csv 로 읽는다 (시트 공유가 '링크 있는 모든 사용자'여야 함)."""
import csv, io, re, urllib.request
from datetime import date

CAL_SHEET = "1BsMpLOB9mpImX3m54z3PmVMKgqYIaRSrzDtS2qm-nPc"      # 격일 달력시간표 (월별 탭)
DISCORD_SHEET = "1oda6dlO8GKOnvUKHw9hOusTLRUyWvdwa7q2Uk48y3Jw"  # 비대면 초대코드 (개강일별 탭)


def _get(url):
    with urllib.request.urlopen(url, timeout=10) as r:
        return r.read().decode("utf-8")


def sheet_tabs(sheet_id):
    """[(탭이름, gid)] 시트 순서대로."""
    t = _get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/htmlview")
    unesc = lambda s: re.sub(r"\\x([0-9a-fA-F]{2})", lambda m: chr(int(m[1], 16)), s).replace("\\/", "/")
    return [(unesc(n), g) for n, g in re.findall(r'\{name: "(.*?)", pageUrl: .*?gid: "(\d+)"', t)]


def sheet_csv(sheet_id, gid):
    return _get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}")


def parse_discord_csv(text):
    """한 탭 → [{start, room, hour, days, weekend, subject, teacher, code, note}].
    강의장 칸이 비면 위 행 강의장을 이어 씀. 시작시간 두 양식:
      평일 '평일 10월 19시(화목금)' → 탭 제목 날짜 + 19시 + 화목금
      주말/방특 '10.17~11.08 / 12:00' → 그 날짜 + 12시"""
    base, room, out = None, "", []
    for r in csv.reader(io.StringIO(text)):
        r = (r + [""] * 9)[:9]
        if base is None:
            m = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", r[1])
            if m:
                base = date(*map(int, m.groups()))
                weekend_tab = "주말" in r[1]
            continue
        if r[1].strip() == "강의장":
            continue
        if r[1].strip():
            room = r[1].strip()
        when, subj = r[3].strip(), r[4].strip()
        rm = re.match(r"\s*([A-Za-z])", room)
        if not (when and subj and rm):
            continue
        start, hour, days = base, None, ""
        w = re.search(r"(\d{1,2})\.(\d{1,2})\s*~.*?/\s*(\d{1,2}):\d{2}", when)
        if w:
            start, hour = date(base.year, int(w[1]), int(w[2])), int(w[3])
        else:
            h = re.search(r"(\d{1,2})시(?:\(([월화수목금토일]+)\))?", when)
            if h:
                hour, days = int(h[1]), h[2] or ""
        if hour is None:
            continue
        out.append(dict(start=start.isoformat(), room=rm[1].upper(), hour=hour, days=days,
                        weekend=weekend_tab or subj.startswith("주말"), subject=subj,
                        teacher=r[5].strip(), code=r[7].strip(), note=r[8].strip()))
    return out


_DATE_RE = re.compile(r"^(?:(\d{1,2})/)?(\d{1,2})$")


def tab_ym(name):
    """탭 이름 '2026.10월 (…)' / '27년2월 (…)' → (2026, 10). 못 읽으면 None."""
    m = re.search(r"(\d{4})\.\s*(\d{1,2})월", name) or re.search(r"(\d{2})년\s*(\d{1,2})월", name)
    if not m:
        return None
    y = int(m[1])
    return (y + 2000 if y < 100 else y, int(m[2]))


def parse_calendar_csv(text, year):
    """격일 달력시간표 한 탭 → 달력 블록 목록.
    블록 = 제목(헤더 위 1~3행) + '일~토' 헤더 + (날짜행, 라벨행) 반복 + 아래 메모(1~3행).
    → [{title, notes, weeks:[[{d, iso, label}]*7]}]"""
    R = list(csv.reader(io.StringIO(text)))
    W = max((len(r) for r in R), default=0)
    R = [r + [""] * (W - len(r)) for r in R]
    cell = lambda i, j: R[i][j].strip() if 0 <= i < len(R) and 0 <= j < W else ""
    heads = [(i, j - 1) for i in range(len(R)) for j in range(1, W - 5)
             if [cell(i, j + k) for k in range(6)] == list("월화수목금토")]
    near_title = {(i - k, c) for i, c in heads for k in (1, 2, 3)}
    is_date_row = lambda i, c: sum(bool(_DATE_RE.match(cell(i, j))) for j in range(c, c + 7)) >= 5
    blocks = []
    for hi, c in heads:
        txt = lambda i: [cell(i, j) for j in range(c, c + 7)]
        title = next(("\n".join(t for t in txt(i) if t) for i in (hi - 1, hi - 2, hi - 3) if i >= 0 and any(txt(i))), "")
        weeks, i, y, mon, prev = [], hi + 1, year, None, 0
        # 월 추적 기준: 블록 안 첫 'M/D' 표기. 그 앞의 숫자는 기준일보다 크면 전달, 작으면 같은 달.
        first = next((m for k in range(hi + 1, min(hi + 20, len(R))) if is_date_row(k, c)
                      for m in map(_DATE_RE.match, txt(k)) if m and m[1]), None)
        while i < len(R) and is_date_row(i, c):
            labels = txt(i + 1) if not is_date_row(i + 1, c) else [""] * 7
            raw = txt(i)
            week = []
            for k, d in enumerate(raw):
                m, iso = _DATE_RE.match(d), ""
                if m:
                    dd = int(m[2])
                    if m[1]:
                        nm = int(m[1])
                        if mon and nm < mon:
                            y += 1
                        mon = nm
                    elif mon is None and first:
                        fm, fd = int(first[1]), int(first[2])
                        mon = fm if dd < fd else (fm - 1 or 12)
                        if fm == 1 and mon == 12:
                            y -= 1
                    elif mon and dd < prev:
                        mon, y = (1, y + 1) if mon == 12 else (mon + 1, y)
                    if mon:
                        iso, prev = f"{y}-{mon:02d}-{dd:02d}", dd
                week.append({"d": d, "iso": iso, "label": labels[k]})
            weeks.append(week)
            i += 2
            if i < len(R) and not any(txt(i)) and is_date_row(i + 1, c):   # 주 사이 빈 행 1개
                i += 1
        if not weeks:
            continue
        notes = []
        for k in range(i, min(i + 4, len(R))):
            if any(h == k and hc == c for h, hc in heads) or (k, c) in near_title:
                break
            notes += [t for t in txt(k) if t and not _DATE_RE.match(t)]
        blocks.append({"title": title, "notes": "\n".join(notes), "weeks": weeks})
    return blocks


def gyeokil_blocks(around, months_before=1, months_after=2):
    """around=(연,월) 기준 앞뒤 달 탭의 달력 블록 전부 (탭 이름 붙여서)."""
    out = []
    base = around[0] * 12 + around[1]
    for name, gid in sheet_tabs(CAL_SHEET):
        ym = tab_ym(name)
        if ym and -months_before <= ym[0] * 12 + ym[1] - base <= months_after:
            for b in parse_calendar_csv(sheet_csv(CAL_SHEET, gid), ym[0]):
                out.append({**b, "tab": name})
    return out


def discord_codes(last_tabs=8):
    """최근 탭 N개의 초대코드 전부. ponytail: 탭이 계속 쌓이므로 최근 N개만 — 오래된 개강분은 필요 없음."""
    rows = []
    for _, gid in sheet_tabs(DISCORD_SHEET)[-last_tabs:]:
        rows += parse_discord_csv(sheet_csv(DISCORD_SHEET, gid))
    return rows


if __name__ == "__main__":
    weekday = (',2026년 10월 19일 평일 정규 비대면수업 초대코드,,,,,,,,\n'
               ',강의장,디스코드 아이디,강의시작 시간,과정명,담당강사,서버명,초대코드,비고\n'
               ',A강의장,x,평일 10월 19시(화목금),일러스트,정아진,s,https://discord.gg/AAA,\n'
               ',,,평일 10월 09시,포토샵,강전욱,s,https://discord.gg/BBB,\n'
               ',,,,,,,,"컴기초 -> 비대면 초대코드 미생성"\n')
    a, b = parse_discord_csv(weekday)
    assert (a["room"], a["hour"], a["days"], a["start"], a["weekend"]) == ("A", 19, "화목금", "2026-10-19", False)
    assert (b["room"], b["hour"], b["code"]) == ("A", 9, "https://discord.gg/BBB")
    weekend = (',2026년 10월 17일 주말 정규 비대면수업 초대코드,,,,,,,,\n'
               ',H 강의장,x,10.17~11.08 / 14:00,주말 / 전응건,박낙규,s,https://discord.gg/CCC,4H\n')
    [c] = parse_discord_csv(weekend)
    assert (c["room"], c["hour"], c["start"], c["weekend"], c["note"]) == ("H", 14, "2026-10-17", True, "4H")

    cal = (',월수금 : 10월 19일~11월 11일,,,,,,,,AI에이전트\n'
           ',,,,,,,,,\n'
           ',일,월,화,수,목,금,토,,\n'
           ',25,26,27,28,29,30,31,,\n'
           ',,월수,,월수,,월수,,,\n'
           ',,,,,,,,,\n'
           ',11/1,2,3,4,5,6,7,,\n'
           ',,월수,,휴강,,,,,\n'
           ',,,,,,,,,\n'
           ',신지아 강사 휴보강,,,,,,,,\n')
    [k] = parse_calendar_csv(cal, 2026)
    assert k["title"] == "월수금 : 10월 19일~11월 11일" and k["notes"] == "신지아 강사 휴보강"
    assert [x["iso"] for x in k["weeks"][0]][:2] == ["2026-10-25", "2026-10-26"]   # '11/1' 앞 숫자 → 전달
    assert k["weeks"][1][3] == {"d": "4", "iso": "2026-11-04", "label": "휴강"}
    assert tab_ym("2026.10월 (10.19개강11.12종강)") == (2026, 10) and tab_ym("27년2월 (개 2.15)") == (2027, 2)
    print("ok")
