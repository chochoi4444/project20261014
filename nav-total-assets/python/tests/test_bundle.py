"""핵심 로직 단일 파일(tools/bundle_core.py)이 패키지와 같은 결과를 내는지"""
import importlib.util
import sys
from pathlib import Path

from make_fixture import make
from nav_total_assets.pipeline import run

ROOT = Path(__file__).resolve().parents[1]


def test_bundle_matches_package(tmp_path):
    sys.path.insert(0, str(ROOT / "tools"))
    import bundle_core
    out = tmp_path / "core.py"
    out.write_text(bundle_core.build(), encoding="utf-8")
    spec = importlib.util.spec_from_file_location("core_bundle", out)
    core = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(core)
    d = make(tmp_path / "fx")
    _, _, a, base_a = core.run(d / "대상.xlsx", [d / "장부"])
    _, _, b, base_b = run(d / "대상.xlsx", [d / "장부"])
    key = lambda r: (r.excel_row, r.status, r.include, round(r.score, 9), r.ledger.value if r.ledger else None)
    assert base_a == base_b and [key(r) for r in a] == [key(r) for r in b]
