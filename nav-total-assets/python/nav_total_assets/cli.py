"""명령줄 실행: python -m nav_total_assets 대상.xlsx 장부폴더 [옵션]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .pipeline import STATUS_KO, load_mapping, run, save_report, save_workbook


def _fmt(n):
    return "" if n is None else f"{round(n):,}"


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        prog="python -m nav_total_assets",
        description="기준가격대장(Excel·PDF)에서 총자산을 찾아 대상 엑셀의 총자산 칸에 입력합니다. 원본은 덮어쓰지 않습니다.",
    )
    p.add_argument("target", help="총자산을 채울 대상 엑셀 (.xlsx)")
    p.add_argument("ledgers", nargs="+", help="기준가격대장 파일 또는 폴더 (여러 개 가능)")
    p.add_argument("-o", "--out", help="저장할 엑셀 경로 (기본: 대상이름_총자산입력.xlsx)")
    p.add_argument("--report", help="매칭 결과 CSV 경로 (기본: 대상이름_매칭결과.csv)")
    p.add_argument("--base-date", help="기준일 YYYY-MM-DD (기본: 대상 엑셀 머리글에서 읽음). 장부 기준일이 다르면 '검토'")
    p.add_argument("--mapping", help="펀드명 매핑 JSON (웹에서 내보낸 파일 또는 {\"종목명\": \"장부 펀드명/파일명\"})")
    p.add_argument("--include-review", action="store_true", help="'검토' 행도 입력 (기본은 '자동'·'수동'만)")
    p.add_argument("--overwrite", action="store_true", help="이미 값이 있는 칸도 덮어쓰기")
    p.add_argument("--dry-run", action="store_true", help="결과만 보여 주고 엑셀은 저장하지 않음")
    p.add_argument("-q", "--quiet", action="store_true", help="행별 결과 출력 생략")
    a = p.parse_args(argv)

    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass

    target_path = Path(a.target)
    if not target_path.is_file():
        p.error(f"대상 엑셀이 없습니다: {target_path}")
    try:
        mapping = load_mapping(a.mapping)
    except Exception as e:
        p.error(f"매핑 파일을 읽지 못했습니다: {e}")

    def progress(i, n, f):
        if not a.quiet:
            print(f"\r장부 읽는 중… {i}/{n}", end="", file=sys.stderr, flush=True)

    try:
        target, ledgers, results, base = run(target_path, a.ledgers, a.base_date, a.overwrite, a.include_review, mapping, progress)
    except ValueError as e:
        print(f"\n오류: {e}", file=sys.stderr)
        return 2
    print("", file=sys.stderr)

    print(f"대상: {target.path.name} · 시트 '{target.sheet_name}' · {target.asset_col_letter}열 '{target.asset_header}' · 종목 {len(target.rows)}행")
    print(f"기준일: {base or '(없음)'} · 장부 {len(ledgers)}개" + (f" (읽기 실패 {sum(1 for l in ledgers if l.error)}개)" if any(l.error for l in ledgers) else ""))
    for l in ledgers:
        if l.error:
            print(f"  ! {l.file_name}: {l.error}")

    if not a.quiet:
        print("\n행   상태    유사도  총자산(원)          검증                  종목명  ←  장부")
        for r in results:
            if r.status in ("nomatch", "skip"):
                continue
            L = r.ledger
            mark = "●" if r.include else " "
            print(f"{r.excel_row:<4} {mark}{STATUS_KO[r.status]:<5} {r.score:>5.2f}  {_fmt(L.value if L else None):>18}  "
                  f"{(L.check[1] if L else ''):<20}  {r.name}  ←  {(L.name or L.file_title) if L else '-'}"
                  + (f"   [{' · '.join(r.notes)}]" if r.notes else ""))
        print("(● = 엑셀에 입력)")

    cnt = {k: sum(1 for r in results if r.status == k) for k in STATUS_KO}
    inc = sum(1 for r in results if r.include)
    print("\n요약: " + " · ".join(f"{STATUS_KO[k]} {v}" for k, v in cnt.items()) + f" · 입력 {inc}행")
    used = {r.ledger.id for r in results if r.ledger}
    unused = [l for l in ledgers if not l.error and l.id not in used]
    if unused:
        print(f"매칭 안 된 장부 {len(unused)}개: " + ", ".join(l.file_name for l in unused[:10]) + (" …" if len(unused) > 10 else ""))

    stem = target.path.with_suffix("")
    report = Path(a.report) if a.report else Path(f"{stem}_매칭결과.csv")
    save_report(results, report)
    print(f"매칭 결과: {report}")
    if a.dry_run:
        print("--dry-run: 엑셀은 저장하지 않았습니다.")
        return 0
    if inc == 0:
        print("입력할 행이 없어 엑셀은 저장하지 않았습니다. '검토' 행을 확인한 뒤 --include-review 또는 --mapping을 쓰세요.")
        return 0
    out = Path(a.out) if a.out else Path(f"{stem}_총자산입력.xlsx")
    n = save_workbook(target, results, out)
    print(f"저장: {out} (총자산 {n}칸)")
    return 0
