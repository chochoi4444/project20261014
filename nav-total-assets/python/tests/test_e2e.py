"""실제 파일로 끝까지 (선택): TARGET=대상.xlsx LEDGERS=장부폴더 [BASE_DATE=YYYY-MM-DD] pytest tests/test_e2e.py -s"""
import os
import zipfile
from pathlib import Path

import pytest

from nav_total_assets.pipeline import run, save_report, save_workbook

pytestmark = pytest.mark.skipif(not (os.environ.get("TARGET") and os.environ.get("LEDGERS")), reason="TARGET/LEDGERS 미지정")


def test_end_to_end(tmp_path):
    target, ledgers, results, base = run(os.environ["TARGET"], [os.environ["LEDGERS"]], os.environ.get("BASE_DATE"))
    for l in ledgers:
        print(f"{l.file_name} | {l.name or l.file_title} | {l.date} | {l.value} | {l.method} | {l.check[1]}")
    out = tmp_path / "out.xlsx"
    n = save_workbook(target, results, out)
    save_report(results, tmp_path / "report.csv")
    assert all(not l.error for l in ledgers)
    # 원본과 비교해 바뀐 부분은 시트 XML과 workbook.xml뿐이어야 한다
    with zipfile.ZipFile(os.environ["TARGET"]) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        changed = [x for x in a.namelist() if a.read(x) != b.read(x)]
        assert set(changed) <= {"xl/workbook.xml"} | {x for x in changed if x.startswith("xl/worksheets/")}
    print(f"입력 {n}칸 → {out}")
