"""펀드명 매칭 회귀 테스트 (가상 펀드명, 웹 버전 tests/matching.test.mjs와 같은 케이스)"""
import json
import os
from pathlib import Path

import pytest

from nav_total_assets.engine import Idf, cosine_sim, match_names, match_status, norm_name, REVIEW_SCORE
from nav_total_assets.pipeline import load_mapping

KO = {"auto": "자동", "review": "검토", "nomatch": "미매칭"}


def ledger(lid, name, manager=""):
    return {"id": lid, "name_cands": [name], "name": name, "manager": manager}


def test_variants_pick_right_fund(cases):
    T = cases["targets"]
    for v in cases["variants"]:
        scores, assign = match_names(T, [ledger(0, v["ledger"], v["manager"])])
        ranked = sorted(T, key=lambda n: -scores[n][0]["score"])
        assert ranked[0] == v["target"], f'{v["how"]}: 1순위 {ranked[0]}'
        assert assign.get(v["target"]) == 0, f'{v["how"]}: 배정 안 됨'
        st = match_status(v["target"], 0, scores)
        assert st in v["allowed"], f'{v["how"]}: 상태 {KO[st]}'


def test_number_pairs_are_not_candidates(cases):
    T = cases["targets"]
    for p in cases["number_pairs"]:
        s = match_names(T, [ledger(0, p["ledger"])])[0][p["target"]][0]["score"]
        assert s < REVIEW_SCORE, f'{p["how"]}: 점수 {s:.2f}'


def test_missing_ledger_not_attached_to_sibling(cases):
    other = next(p["ledger"] for p in cases["number_pairs"] if p["how"].startswith("89호"))
    _, assign = match_names(cases["targets"], [ledger(0, other)])
    assert 0 not in assign.values()


def test_all_ledgers_one_to_one(cases):
    T = cases["targets"]
    ls = [ledger(i, v["ledger"], v["manager"]) for i, v in enumerate(cases["variants"])]
    ls += [ledger(100 + i, p["ledger"]) for i, p in enumerate(cases["number_pairs"])
           if not any(norm_name(t) == norm_name(p["ledger"]) for t in T)]
    _, assign = match_names(T, ls)
    for i, v in enumerate(cases["variants"]):
        assert assign.get(v["target"]) == i
    used = [x for x in assign.values() if x is not None]
    assert len(used) == len(set(used))


def test_cosine_properties():
    assert cosine_sim("가온블랙락73", "가온블랙락73") == pytest.approx(1)
    assert cosine_sim("가온블랙락", "누리USDL") == 0
    idf = Idf(["가온부동산", "누리부동산", "다온부동산", "가온인프라"])
    ratio = lambda a, b: cosine_sim(a, b, idf) / cosine_sim(a, b)
    assert ratio("가온부동산", "누리부동산") < ratio("가온부동산", "가온인프라")


def test_forced_mapping_wins(cases):
    T = cases["targets"]
    ls = [ledger(0, "전혀 다른 이름의 장부"), ledger(1, cases["variants"][0]["ledger"])]
    _, assign = match_names(T, ls, {T[0]: 0})
    assert assign[T[0]] == 0


def test_load_mapping_web_and_plain(tmp_path):
    web = tmp_path / "web.json"
    web.write_text(json.dumps({"version": 1, "mapping": {norm_name("가온 101호"): norm_name("가온101호 장부"), "X": "__none__"}}, ensure_ascii=False), encoding="utf-8")
    plain = tmp_path / "plain.json"
    plain.write_text(json.dumps({"가온 101호": "가온101호 장부", "X": None}, ensure_ascii=False), encoding="utf-8")
    assert load_mapping(web) == load_mapping(plain)


@pytest.mark.skipif(not os.environ.get("MATCH_CASES"), reason="MATCH_CASES 미지정")
def test_real_cases():
    """실제 종목 케이스 (저장소 밖 JSON, 웹 버전과 같은 형식)"""
    data = json.loads(Path(os.environ["MATCH_CASES"]).read_text(encoding="utf-8"))
    for c in data["cases"]:
        target, variant = c[0], c[1]
        manager = c[2] if len(c) > 2 else ""
        allowed = c[3] if len(c) > 3 else ["auto", "review"]
        scores, _ = match_names(data["targets"], [ledger(0, variant, manager)])
        ranked = sorted(data["targets"], key=lambda n: -scores[n][0]["score"])
        assert ranked[0] == target, f"{variant} → {ranked[0]}"
        assert match_status(target, 0, scores) in allowed
