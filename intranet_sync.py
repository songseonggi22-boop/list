"""뉴인트라넷(kegone) 강의시간표 동기화 — 운영관리 > 시간표관리 > 시간표관리 '기간' 보기의 [엑셀다운로드]와 같은 파일을 받는다.
화면 JS가 쓰는 API를 그대로 호출:
  POST /api/auth/login                     {userId, password} → data.user.accessToken
  POST /api/op/b08/b0801/scheduler/excel   {siteCode, place, stDate, edDate, view:'custom', …} → 기간시간표 xlsx
아이디/비번은 Streamlit Secrets(INTRANET_ID / INTRANET_PW)에서만 읽는다 — 코드·DB·git 에 남기지 않음."""
import calendar, json, urllib.request
from datetime import date

API = "https://kegone.koreagroupware.com/api"


def _post(path, body, token=None):
    req = urllib.request.Request(API + path, data=json.dumps(body).encode(), method="POST",
                                 headers={"Content-Type": "application/json",
                                          **({"Authorization": f"Bearer {token}"} if token else {})})
    return urllib.request.urlopen(req, timeout=60)


def _pick(d, *keys):
    return next((d[k] for k in keys if isinstance(d, dict) and d.get(k) not in (None, "")), "")


def login(user_id, password):
    """→ (token, places) — places = [{siteCode, place, label}] 로그인 계정이 볼 수 있는 지점."""
    with _post("/auth/login", {"userId": user_id, "password": password}) as r:
        res = json.load(r)
    code = (res.get("errorCode") or {}).get("code")
    if code not in ("LOGIN_SUCCESS", "OK"):
        raise RuntimeError("인트라넷 로그인 실패: " + str((res.get("errorCode") or {}).get("message") or code))
    data = res.get("data") or {}
    token = (data.get("user") or {}).get("accessToken")
    if not token:
        raise RuntimeError("인트라넷 로그인 응답에 토큰이 없어요")
    places = [{"siteCode": str(_pick(p, "siteCode", "site_code")),
               "place": str(_pick(p, "placeNo", "place_no", "place")),
               "label": str(_pick(p, "placeName", "place_name", "name", "label"))}
              for p in data.get("placeList") or []]
    return token, [p for p in places if p["siteCode"] and p["place"]]


def month_ranges(start, months):
    """start(date)가 속한 달부터 months개 달의 (1일, 말일)."""
    y, m = start.year, start.month
    for _ in range(months):
        yield date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def period_excel(token, site_code, place, st, ed, weekend_yn):
    """기간시간표 xlsx → (파일명, bytes). weekend_yn 'N'=평일 / 'Y'=주말 — 도구가 (월, 평일/주말)마다 파일 1개라 나눠 받는다.
    파일명은 대시보드 파서가 월을 판정하는 '기간시간표_YYYYMMDD_YYYYMMDD' 꼴로 맞춘다."""
    body = {"siteCode": site_code, "place": place, "placeNo": place, "subjectNo": "", "subjectName": "",
            "classroomNo": "", "weekendYn": weekend_yn, "capacityStatus": "",
            "stDate": st.isoformat(), "edDate": ed.isoformat(), "view": "custom", "daySelect": "ALL"}
    with _post("/op/b08/b0801/scheduler/excel", body, token) as r:
        raw = r.read()
        cd = r.headers.get("Content-Disposition", "")
    if not raw.startswith(b"PK"):   # xlsx(zip) 아니면 오류 JSON
        raise RuntimeError("시간표 엑셀을 못 받았어요: " + raw[:200].decode("utf-8", "replace"))
    return f"기간시간표_{st:%Y%m%d}_{ed:%Y%m%d}_{'주말' if weekend_yn == 'Y' else '평일'}.xlsx", raw


def sync(user_id, password, start, months=3, place_filter="대전"):
    """→ ([(파일명, bytes)], [지점 라벨]). 지점 이름에 place_filter 가 든 곳만 (없으면 전부)."""
    token, places = login(user_id, password)
    targets = [p for p in places if place_filter in p["label"]] or places
    if not targets:
        raise RuntimeError("계정에 연결된 지점(placeList)이 없어요")
    files = []
    for p in targets:
        for st, ed in month_ranges(start, months):
            for yn in ("N", "Y"):
                name, raw = period_excel(token, p["siteCode"], p["place"], st, ed, yn)
                files.append((f"[{p['label'] or p['place']}]{name}", raw))
    return files, [p["label"] for p in targets]


if __name__ == "__main__":
    assert list(month_ranges(date(2026, 11, 14), 3)) == [
        (date(2026, 11, 1), date(2026, 11, 30)), (date(2026, 12, 1), date(2026, 12, 31)),
        (date(2027, 1, 1), date(2027, 1, 31))]
    print("ok")
