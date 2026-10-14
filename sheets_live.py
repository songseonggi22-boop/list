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


def sheet_view_url(sheet_id, gid):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/htmlview/sheet?headers=false&gid={gid}"


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
    print("ok")
