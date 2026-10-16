"""펀드명 매칭 엔진.

웹 버전(src/app.html의 '매칭 엔진' 구역)과 같은 규칙·같은 숫자를 낸다.
tests/test_parity.py가 두 구현의 점수를 대조한다.
"""
from __future__ import annotations

import math
import re
import unicodedata
from typing import Iterable

# ---------- 텍스트 정규화 ----------

def nfkc(s) -> str:
    return unicodedata.normalize("NFKC", "" if s is None else str(s))


_LABEL_DROP = re.compile(r"[\s\[\]▶:：·]")


def norm_label(s) -> str:
    return _LABEL_DROP.sub("", nfkc(s))


ALIASES = [
    (re.compile(r"(^|[^A-Z])UP(?=[^A-Z]|$)"), r"\1유나이티드파트너스"),
    (re.compile(r"BLACKROCK"), "블랙락"), (re.compile(r"블랙록"), "블랙락"),
    (re.compile(r"SAMSUNG"), "삼성"), (re.compile(r"SHINHAN"), "신한"), (re.compile(r"KIWOOM"), "키움"),
    (re.compile(r"HANA(?![A-Z])"), "하나"),
    (re.compile(r"IGIS"), "이지스"), (re.compile(r"HYUNDAI"), "현대"), (re.compile(r"KYOBO"), "교보"),
    (re.compile(r"MIRAE"), "미래에셋"),
]
_NAME_DROP = re.compile(r"[^가-힣A-Z0-9]")


def norm_name(s) -> str:
    t = nfkc(s).upper()
    for rx, rep in ALIASES:
        t = rx.sub(rep, t)
    return _NAME_DROP.sub("", t)


STOP = ["전문투자형사모", "전문투자형", "일반사모", "전문사모", "사모", "공모", "투자신탁", "투자회사", "투자조합",
        "신탁", "투자", "펀드", "재간접형", "재간접", "증권", "자투자", "모투자", "유한회사", "주식회사", "전문", "일반",
        "CLASS", "클래스"]


def core_name(s) -> str:
    t = norm_name(s)
    for w in STOP:
        t = t.replace(w, "")
    t = re.sub(r"제(?=[0-9])", "", t)
    return re.sub(r"([0-9])(호|종)", r"\1", t)


def nums(s) -> list[str]:
    found = re.findall(r"[0-9]+", nfkc(s))
    # 5자리 이상은 펀드코드·날짜로 보고 호수 비교에서 뺀다
    return sorted({str(int(x)) for x in found if len(x) < 5})


# ---------- 유사도 ----------
# 코사인 유사도: 핵심 이름을 글자 1~3개 묶음(n-gram) 벡터로 바꾸고, 여러 펀드명에 흔한 묶음은
# TF-IDF로 가중치를 낮춰 그 펀드에만 있는 글자가 점수를 좌우하게 한다.
GRAM_SIZES = (1, 2, 3)
_gram_cache: dict[str, dict[str, int]] = {}


def grams(s: str) -> dict[str, int]:
    m = _gram_cache.get(s)
    if m is not None:
        return m
    m = {}
    for n in GRAM_SIZES:
        for i in range(0, len(s) - n + 1):
            g = s[i:i + n]
            m[g] = m.get(g, 0) + 1
    if len(_gram_cache) > 20000:
        _gram_cache.clear()
    _gram_cache[s] = m
    return m


class Idf:
    """비교 대상 전체(대상 종목명 + 장부 펀드명)의 핵심 이름으로 만든 IDF"""

    def __init__(self, core_names: Iterable[str]):
        uniq = list(dict.fromkeys(core_names))
        df: dict[str, int] = {}
        for s in uniq:
            for g in grams(s):
                df[g] = df.get(g, 0) + 1
        n = len(uniq)
        self.idf = {g: math.log((1 + n) / (1 + d)) + 1 for g, d in df.items()}
        self.unseen = math.log(1 + n) + 1

    def weight(self, g: str) -> float:
        return self.idf.get(g, self.unseen)


def cosine_sim(a: str, b: str, idf: Idf | None = None) -> float:
    if not a or not b:
        return 0.0
    A, B = grams(a), grams(b)
    w = idf.weight if idf else (lambda g: 1)
    dot = na = nb = 0.0
    for g, c in A.items():
        x = c * w(g)
        na += x * x
        d = B.get(g)
        if d:
            dot += x * d * w(g)
    for g, c in B.items():
        y = c * w(g)
        nb += y * y
    return dot / math.sqrt(na * nb) if na and nb else 0.0


def lcs_ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    prev = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        cur = [0]
        for j in range(1, len(b) + 1):
            cur.append(prev[j - 1] + 1 if a[i - 1] == b[j - 1] else max(prev[j], cur[j - 1]))
        prev = cur
    return 2 * prev[len(b)] / (len(a) + len(b))


# 펀드명 맨 앞의 사명(운용사). 펀드명은 사명으로 시작하므로 사명이 다르면 다른 펀드다.
# norm_name을 거친 모양으로 적는다 (영문 사명은 ALIASES가 한글로 바꿔 준다)
MANAGERS = ["KB", "NH아문디", "NH", "신한", "삼성", "미래에셋", "한화", "키움", "이지스", "마스턴", "코람코", "블랙락",
            "유나이티드파트너스", "교보", "현대", "한국투자", "IBK", "DB", "우리", "하나", "하나대체", "하나대체투자", "대신", "메리츠",
            "흥국", "타임폴리오", "트러스톤", "신영", "베어링", "슈로더", "피델리티", "JP모간", "멀티에셋", "브이아이", "에셋플러스",
            "유진", "BNK", "DGB", "하이", "KTB", "다올", "HDC", "칸서스", "캡스톤", "라살", "인베스코", "알리안츠", "맥쿼리",
            "파인스트리트", "제이알", "케이리츠", "한국투자리얼에셋", "삼성SRA", "이지스자산", "코람코자산"]
_MGR_SUFFIX = re.compile(r"(자산운용|투자운용|자산관리|운용|ASSETMANAGEMENT|MANAGEMENT|INVESTMENTTRUST|INVESTMENT|TRUST|MGT|ASSET|COLTD|CO|LTD|주식회사)+$")


def mgr_prefix(s) -> str:
    # '(블랙락73호)'처럼 괄호로 시작하면 별칭·메모라 펀드명 앞머리로 보지 않는다
    if re.match(r"\s*[(\[{（【<]", nfkc(s)):
        return ""
    n = norm_name(s)
    best = ""
    for m in MANAGERS:
        if len(m) > len(best) and n.startswith(m):
            best = m
    return best


def _mgr_compat(a: str, b: str) -> bool:
    return a.startswith(b) or b.startswith(a)


def header_mgr(manager) -> str:
    """장부 머리글 운용사명에서 '자산운용' 같은 꼬리를 뗀 사명"""
    t = norm_name(manager)
    while t:
        u = _MGR_SUFFIX.sub("", t)
        if u == t:
            break
        t = u
    return t if len(t) >= 2 else ""


# 이름 유사도 = 코사인(TF-IDF) 60% + 글자 순서 일치(LCS) 40%
COSINE_WEIGHT = 0.6


def score_name(target: str, ledger: dict, idf: Idf | None = None) -> tuple[float, str]:
    """대상 종목명과 장부 하나의 유사도 (0~1)와 근거.

    ledger: {"name_cands": [...], "manager": "..."}
    """
    tn, tc, tnum = norm_name(target), core_name(target), nums(target)
    best, why = 0.0, ""
    for cand in ledger.get("name_cands", []):
        if not cand:
            continue
        if norm_name(cand) == tn:
            return 1.0, "펀드명 완전 일치"
        cc, cnum = core_name(cand), nums(cand)
        cs = cosine_sim(tc, cc, idf)
        s = COSINE_WEIGHT * cs + (1 - COSINE_WEIGHT) * lcs_ratio(tc, cc)
        w = f"코사인 {cs:.2f}"
        if tnum and cnum:
            # 호수·종류(1종/2종, 67호/67-1호, 3-1호/3-3호)가 하나라도 다르면 다른 펀드로 본다
            if tnum == cnum:
                s = s * 0.85 + 0.15
                w += " · 호수 일치"
            else:
                s *= 0.35
                w += " · 호수·종류 다름" if any(x in cnum for x in tnum) else " · 호수 다름"
        elif len(tnum) != len(cnum):
            s *= 0.8
        if s > best:
            best, why = s, w
    # 사명: 장부 쪽(펀드명·파일명 앞머리, 머리글 운용사)과 대상 종목명 앞머리를 맞춘다. 호수보다 우선
    tm, hdr = mgr_prefix(target), header_mgr(ledger.get("manager") or "")
    lm = [mgr_prefix(c) for c in ledger.get("name_cands", []) if c]
    if hdr:
        lm.append(mgr_prefix(hdr) or hdr)
    lms = [x for x in lm if x]
    if (hdr and tn.startswith(hdr)) or (tm and any(_mgr_compat(tm, x) for x in lms)):
        return min(1.0, best + 0.12), why + " · 사명 일치"
    if tm and lms:
        return best * 0.3, why + f" · 사명 다름({tm}≠{lms[0]})"
    mgr = norm_name(ledger.get("manager") or "")
    if mgr and len(tn) >= 4:
        if tn[:4] in mgr:
            best += 0.1
            why += " · 운용사 일치"
        elif re.fullmatch(r"[가-힣]{2}", tn[:2]) and tn[:2] in mgr:
            best += 0.04
    return min(1.0, best), why


AUTO_SCORE, AUTO_MARGIN, REVIEW_SCORE = 0.85, 0.08, 0.55
# 코사인은 같은 펀드와 다른 펀드의 점수 차이가 크게 벌어진다: 점수가 조금 낮아도 2순위와 격차가 크면 자동
AUTO_CLEAR_SCORE, AUTO_CLEAR_MARGIN = 0.75, 0.4


def match_names(names: Iterable[str], ledgers: list[dict], forced: dict | None = None):
    """대상 종목명들과 장부들을 1:1로 짝짓는다.

    ledgers: [{"id", "name_cands", "manager", "name"?, "file_title"?}]
    forced:  {종목명: 장부 id 또는 None}
    반환: (scores, assign)
      scores = {종목명: [{"id", "score", "why"}, ...]}  (장부 순서대로)
      assign = {종목명: 장부 id 또는 None}
    """
    forced = forced or {}
    uniq = list(dict.fromkeys(names))
    idf = Idf([core_name(x) for x in uniq + [c for l in ledgers for c in l.get("name_cands", []) if c]])
    scores = {}
    for n in uniq:
        lst = []
        for l in ledgers:
            sc, why = score_name(n, l, idf)
            lst.append({"id": l["id"], "score": sc, "why": why})
        scores[n] = lst
    # 호수 유일 일치: 장부 이름이 깨져도 두 자리 이상 호수가 대상 하나와만 맞으면 검토 후보로 올림
    for l in ledgers:
        cn = nums(l.get("name") or l.get("file_title") or (l.get("name_cands") or [""])[0] or "")
        if not any(len(x) >= 2 for x in cn):
            continue
        hits = [n for n in uniq if nums(n) == cn]
        if len(hits) == 1:
            e = next(x for x in scores[hits[0]] if x["id"] == l["id"])
            if e["score"] < 0.62 and "사명 다름" not in e["why"]:
                e["score"], e["why"] = 0.62, "호수만 유일하게 일치"
    assign, used = {}, set()
    for n, lid in forced.items():
        if n in scores:
            assign[n] = lid
            if lid is not None:
                used.add(lid)
    pairs = [(n, e) for n, lst in scores.items() if n not in assign for e in lst if e["score"] >= REVIEW_SCORE]
    pairs.sort(key=lambda p: -p[1]["score"])  # 안정 정렬: 웹 버전과 같은 순서
    for n, e in pairs:
        if n in assign or e["id"] in used:
            continue
        assign[n] = e["id"]
        used.add(e["id"])
    return scores, assign


def match_status(name: str, lid, scores: dict) -> str:
    """배정된 장부의 이름 점수만으로 본 상태: auto / review / nomatch"""
    lst = scores.get(name, [])
    e = next((x for x in lst if x["id"] == lid), None)
    if e is None:
        return "nomatch"
    second = max([x["score"] for x in lst if x["id"] != lid] + [0])
    margin = e["score"] - second
    if e["score"] >= AUTO_SCORE and margin >= AUTO_MARGIN:
        return "auto"
    if e["score"] >= AUTO_CLEAR_SCORE and margin >= AUTO_CLEAR_MARGIN:
        return "auto"
    return "review"
