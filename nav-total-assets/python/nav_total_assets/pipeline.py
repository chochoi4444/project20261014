"""대상 엑셀 + 장부들 → 행별 매칭 결과 → 저장. 웹 버전의 runMatching과 같은 판정이다."""
from __future__ import annotations

import csv
import json
from dataclasses import dataclass, field
from pathlib import Path

from .engine import match_names, match_status, norm_name, AUTO_SCORE
from .ledger import Ledger, parse_ledgers
from .workbook import Target, has_number, is_text_mark, parse_target, write_values

LEDGER_EXTS = {".xlsx", ".xlsm", ".xls", ".csv", ".htm", ".html"}
STATUS_KO = {"auto": "자동", "review": "검토", "manual": "수동", "nomatch": "미매칭", "skip": "제외"}


@dataclass
class RowResult:
    excel_row: int
    name: str
    code: str
    cls: str
    existing: object
    status: str = "nomatch"
    ledger: Ledger | None = None
    score: float = 0.0
    why: str = ""
    include: bool = False
    notes: list = field(default_factory=list)


def collect_ledger_files(paths) -> list[Path]:
    out = []
    for p in map(Path, paths):
        if p.is_dir():
            out += sorted(x for x in p.rglob("*") if x.is_file() and x.suffix.lower() in LEDGER_EXTS)
        elif p.is_file():
            out.append(p)
    return [p for p in out if not p.name.startswith("~$")]


def load_mapping(path: Path | None) -> dict:
    """웹에서 내보낸 매핑 JSON({"mapping": {...}})이나 직접 쓴 {"종목명": "장부 펀드명 또는 파일명"}을 읽는다.
    값이 null 또는 "__none__"이면 그 종목은 매칭하지 않는다."""
    if not path:
        return {}
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    m = data.get("mapping", data) if isinstance(data, dict) else {}
    return {norm_name(k): ("__none__" if v in (None, "__none__") else norm_name(v)) for k, v in m.items()}


def _ledger_by_key(ledgers: list[Ledger], key: str) -> Ledger | None:
    for l in ledgers:
        if key in (l.key, norm_name(l.file_title), norm_name(Path(l.file_name).stem)):
            return l
    return None


def run(target_path, ledger_paths, base_date: str | None = None, overwrite: bool = False,
        include_review: bool = False, mapping: dict | None = None, progress=None):
    target: Target = parse_target(target_path)
    files = collect_ledger_files(ledger_paths)
    ledgers = []
    for i, f in enumerate(files):
        if progress:
            progress(i + 1, len(files), f)
        ledgers += parse_ledgers(f, len(ledgers))
    ok = [l for l in ledgers if not l.error]
    base = base_date or target.base_date
    mapping = mapping or {}

    uniq = list(dict.fromkeys(r.name for r in target.rows))
    forced = {}
    for n in uniq:
        key = mapping.get(norm_name(n))
        if not key:
            continue
        if key == "__none__":
            forced[n] = None
            continue
        l = _ledger_by_key(ok, key)
        if l:
            forced[n] = l.id
    scores, assign = match_names(uniq, [l.as_match_input() for l in ok], forced)
    by_id = {l.id: l for l in ok}

    results = []
    for r in target.rows:
        res = RowResult(r.excel_row, r.name, r.code, r.cls, r.existing)
        results.append(res)
        if is_text_mark(r.existing):
            res.status = "skip"
            res.notes.append(f"‘{r.existing.strip()}’ 표시 행")
            continue
        lst = sorted(scores.get(r.name, []), key=lambda x: -x["score"])
        is_forced = r.name in forced
        lid = assign.get(r.name)
        if lid is None:
            res.status = "manual" if is_forced else "nomatch"
            if is_forced:
                res.notes.append("매칭 안 함으로 지정됨")
            elif lst and lst[0]["score"] >= 0.3:
                res.notes.append(f"가장 가까운 후보 {lst[0]['score']:.2f} {by_id[lst[0]['id']].name or by_id[lst[0]['id']].file_title}")
            continue
        L = by_id[lid]
        e = next((x for x in lst if x["id"] == lid), {"score": 1, "why": ""})
        res.ledger = L
        res.score = 1.0 if is_forced else e["score"]
        res.why = "지정된 매핑" if is_forced else e["why"]
        name_status = match_status(r.name, lid, scores)
        date_ok = not base or not L.date or L.date == base
        if not date_ok:
            res.notes.append(f"기준일 {L.date} ≠ {base}")
        if L.check[0] == "warn":
            res.notes.append(L.check[1])
        if L.value is None:
            res.notes.append("총자산 값 없음")
        if is_forced:
            res.status = "manual"
        elif name_status == "auto" and date_ok and L.check[0] != "warn" and L.value is not None:
            res.status = "auto"
        else:
            res.status = "review"
            if name_status != "auto":
                res.notes.insert(0, "유사도 낮음" if e["score"] < AUTO_SCORE else "비슷한 후보 있음")
        res.include = L.value is not None and (res.status == "auto" or (res.status == "manual" and date_ok)
                                               or (include_review and res.status == "review"))
        if has_number(r.existing) and not overwrite and res.include:
            res.include = False
            res.notes.append("기존 값 있음 (덮어쓰기 꺼짐)")
    return target, ledgers, results, base


def date_mismatch(results: list[RowResult], base: str | None):
    """매칭된 장부의 가장 흔한 기준일이 기준일과 다르면 (그 날짜, 해당 행 수), 아니면 (None, 0)"""
    dates: dict[str, int] = {}
    for r in results:
        if r.ledger and r.ledger.date:
            dates[r.ledger.date] = dates.get(r.ledger.date, 0) + 1
    if not dates or not base:
        return None, 0
    common = max(dates.items(), key=lambda kv: kv[1])[0]
    if common == base:
        return None, 0
    return common, sum(1 for r in results if r.ledger and r.ledger.date and r.ledger.date != base)


def save_workbook(target: Target, results: list[RowResult], out: Path) -> int:
    values = {f"{target.asset_col_letter}{r.excel_row}": r.ledger.value for r in results if r.include and r.ledger}
    return write_values(target.path, out, target.sheet_index, values)


def save_report(results: list[RowResult], out: Path):
    head = ["행", "종목명", "표준코드", "상태", "입력", "유사도", "매칭근거", "기준가격대장 파일", "장부 펀드명", "장부 기준일",
            "총자산", "산출근거", "검증", "순자산", "부채", "기준가", "메모", "매핑키(장부)"]
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(head)
        for r in results:
            L = r.ledger
            fv = lambda k: ("" if not L or k not in L.fields else L.fields[k]["v"])
            w.writerow([r.excel_row, r.name, r.code, STATUS_KO[r.status], "Y" if r.include else "",
                        f"{r.score:.3f}" if r.score else "", r.why, (L.file_name + (f" [{L.sheet}]" if L.sheet else "")) if L else "",
                        (L.name or L.file_title) if L else "", (L.date or "") if L else "",
                        round(L.value) if L and L.value is not None else "", L.method if L else "",
                        L.check[1] if L else "", fv("net"), fv("liab"), fv("nav"), " / ".join(r.notes),
                        (L.name or L.file_title) if L else ""])
