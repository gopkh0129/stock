# 스크리닝이 끝난 뒤, 기준일에 VPA 패턴이 새로 생긴 종목을 텔레그램으로 보냅니다.
#  - 대상: 눌림목 후보(사이트 기본 조건) + VPA 탭(점수 0.90 이상)
#  - 매수가/손절가/목표가 계산은 사이트(index.html의 pricePlan)와 같은 규칙
#  - 필요한 GitHub Secrets: TELEGRAM_TOKEN, TELEGRAM_CHAT_ID (없으면 알림만 건너뜀)
import json, os, sys, math, datetime, html, urllib.request, urllib.parse

DATA_PATH = "docs/data/latest.json"
SITE = "https://gopkh0129.github.io/stock/"
VPA_MIN = 0.90
WARN_PATTERNS = {"No Demand"}
STOP_PCT, TARGET_PCT = 7, 21   # 손절 -7% / 목표 +21% (손익비 1:3)


def tick_of(p):
    return 1 if p < 2000 else 5 if p < 5000 else 10 if p < 20000 else 50 if p < 50000 else \
        100 if p < 200000 else 500 if p < 500000 else 1000

def tick_up(p):
    t = tick_of(p); return math.ceil(p / t) * t

def tick_dn(p):
    t = tick_of(p); return math.floor(p / t) * t


def price_plan(s, c):
    if not c or not c.get("d"):
        return None
    n = len(c["d"])
    si = c["d"].index(s["매칭일"]) if s.get("매칭일") in c["d"] else n - 1
    last = c["c"][-1]
    buy = tick_up(c["h"][si] + tick_of(c["h"][si]))
    stop = tick_up(buy * (1 - STOP_PCT / 100))
    target = tick_dn(buy * (1 + TARGET_PCT / 100))
    if s.get("패턴") in WARN_PATTERNS:
        status = "⛔ 주의패턴-매수금지"
    elif last >= buy * 1.03:
        status = "이미 돌파 · 되돌림 대기"
    elif last >= buy:
        status = "✅ 돌파 진행 중 (매수가 +3% 이내만 매수)"
    else:
        status = f"돌파 대기 (현재가보다 +{(buy - last) / last * 100:.1f}% 위)"
    return dict(last=last, buy=buy, stop=stop, target=target,
                risk=(buy - stop) / buy * 100, status=status)


def pick_rows(d):
    cond = d["meta"]["기본조건"]
    leaders = {r["종목코드"] for r in d.get("leaders1m", []) + d.get("leaders2w", [])}
    def ge(v, t): return t is None or (v is not None and v >= t)
    def le(v, t): return t is None or (v is not None and v <= t)
    rows = []
    for s in d["stocks"]:
        if cond.get("VPA_후보") == "리더" and s["종목코드"] not in leaders:
            continue
        gap_lo = cond.get("이격_하한") or 0
        if (ge(s.get("1개월 등락률(%)"), cond.get("최소_1개월")) and ge(s.get("2주 등락률(%)"), cond.get("최소_2주"))
                and ge(s.get("업종RS(1개월)"), cond.get("업종RS_하한"))
                and ge(s.get("종목RS"), cond.get("종목RS_하한")) and le(s.get("종목RS"), cond.get("종목RS_상한"))
                and le(s.get("20일선 이격"), cond.get("이격_상한")) and (not gap_lo or ge(s.get("20일선 이격"), gap_lo))
                and le(s.get("거래량비(3일/20일)"), cond.get("거래량비_상한"))):
            rows.append(s)
    return rows


def fmt(v):
    return f"{int(v):,}"


def block(title, rows, d):
    if not rows:
        return f"<b>[{title}]</b> 오늘 새 패턴 없음"
    rows = sorted(rows, key=lambda s: -(s.get("점수") or 0))
    out = [f"<b>[{title}]</b> 오늘 패턴 {len(rows)}종목"]
    for i, s in enumerate(rows, 1):
        p = price_plan(s, d["candles"].get(s["종목코드"]))
        sig = f" ({s['업종 신호']})" if s.get("업종 신호") else ""
        out.append(f"\n{i}) <b>{html.escape(s['종목명'])}</b> {s['종목코드']} · {html.escape(s.get('업종명') or '')}{html.escape(sig)}")
        out.append(f"   {s.get('패턴')} · 점수 {s.get('점수', 0):.2f}")
        if p:
            out.append(f"   현재가 {fmt(p['last'])} / 매수가 {fmt(p['buy'])}")
            out.append(f"   손절가 {fmt(p['stop'])} (-{p['risk']:.1f}%) / 목표가 {fmt(p['target'])} (+{TARGET_PCT}%)")
            out.append(f"   상태: {p['status']}")
    return "\n".join(out)


def build_message(d):
    day = d["meta"]["기준일"]
    wd = "월화수목금토일"[datetime.date.fromisoformat(day).weekday()]
    today = lambda rows: [s for s in rows if s.get("매칭일") == day]
    pick = today(pick_rows(d))
    vpa = today([s for s in d["stocks"] if (s.get("점수") or 0) >= VPA_MIN])
    parts = [f"📈 <b>스윙스크리너 · {day[5:].replace('-', '/')} ({wd})</b>",
             block("눌림목 후보", pick, d), block(f"VPA {VPA_MIN:.2f} 이상", vpa, d),
             f"🔗 {SITE}"]
    return "\n\n".join(parts)


def send(text):
    token, chat = os.environ.get("TELEGRAM_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat:
        print("텔레그램 설정(Secrets)이 없어 알림을 건너뜁니다.")
        return
    chunks, cur = [], ""
    for line in text.split("\n"):           # 텔레그램 한 메시지 최대 4096자
        if len(cur) + len(line) + 1 > 3800:
            chunks.append(cur); cur = ""
        cur += line + "\n"
    chunks.append(cur)
    for ch in chunks:
        body = urllib.parse.urlencode({"chat_id": chat, "text": ch, "parse_mode": "HTML",
                                       "disable_web_page_preview": "true"}).encode()
        with urllib.request.urlopen(f"https://api.telegram.org/bot{token}/sendMessage", body, timeout=30) as r:
            print("텔레그램 전송:", r.status)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--fail":
        send("⚠️ <b>오늘 스크리닝이 실패했습니다.</b>\n사이트는 마지막 성공 결과를 보여줍니다.\n"
             "GitHub Actions 화면을 확인해 주세요:\nhttps://github.com/gopkh0129/stock/actions")
    else:
        with open(DATA_PATH, encoding="utf-8") as f:
            msg = build_message(json.load(f))
        print(msg)
        send(msg)
