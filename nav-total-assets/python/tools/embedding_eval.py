"""임베딩 모델 + 코사인 유사도로 펀드명 매칭을 시험하고, 지금 방식(글자 n-gram TF-IDF 코사인)과 비교한다.

    python tools/embedding_eval.py --model 모델폴더 [--prefix "query: "] [--cases 실제케이스.json] [--detail]

케이스: ../tests/fixtures/matching-cases.json (가상 펀드명) + --cases (실제 종목, 저장소 밖 파일)
비교 방식
  현재         글자 1~3gram TF-IDF 코사인 60% + 글자 순서 40% + 호수 규칙 + 운용사 가점 (프로그램이 쓰는 방식)
  임베딩       모델 임베딩의 코사인 유사도만 (펀드명 그대로)
  임베딩+규칙  임베딩 코사인에 같은 호수 규칙·운용사 가점 적용
  혼합         현재 50% + 임베딩+규칙 50%
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from nav_total_assets.engine import match_names, norm_name, nums  # noqa: E402
from nav_total_assets.embedding import OnnxEmbedder, cosine_matrix  # noqa: E402

FIXTURE = HERE.parent.parent / "tests" / "fixtures" / "matching-cases.json"


def number_rule(score, target, cand, manager=""):
    """프로그램과 같은 호수 규칙·운용사 가점"""
    tnum, cnum = nums(target), nums(cand)
    if norm_name(target) == norm_name(cand):
        return 1.0
    if tnum and cnum:
        score = score * 0.85 + 0.15 if tnum == cnum else score * 0.35
    elif len(tnum) != len(cnum):
        score *= 0.8
    tn, mg = norm_name(target), norm_name(manager)
    if mg and len(tn) >= 4 and tn[:4] in mg:
        score += 0.1
    return min(1.0, score)


def load_sets(extra):
    fx = json.loads(FIXTURE.read_text(encoding="utf-8"))
    sets = [("가상", fx["targets"], [(v["target"], v["ledger"], v["manager"]) for v in fx["variants"]],
             [(p["target"], p["ledger"]) for p in fx["number_pairs"]])]
    if extra:
        d = json.loads(Path(extra).read_text(encoding="utf-8"))
        sets.append(("실제", d["targets"], [(c[0], c[1], c[2] if len(c) > 2 else "") for c in d["cases"]], d.get("pairs", [])))
    return sets


def evaluate(emb: OnnxEmbedder, sets, detail=False):
    methods = ["현재", "임베딩", "임베딩+규칙", "혼합"]
    stats = {m: {"ok": 0, "n": 0, "min_ok": 1.0, "max_2nd": 0.0, "min_gap": 1.0, "pair_max": 0.0, "fails": []} for m in methods}
    t0 = time.time()
    n_embedded = 0
    for set_name, targets, cases, pairs in sets:
        tv = emb.embed(targets)
        n_embedded += len(targets)
        for target, ledger, manager in cases + [(t, l, "") for t, l in pairs]:
            is_pair = (target, ledger) in [tuple(p) for p in pairs]
            lv = emb.embed([ledger])
            n_embedded += 1
            e = cosine_matrix(lv, tv)[0]
            cur_scores, _ = match_names(targets, [{"id": 0, "name_cands": [ledger], "name": ledger, "manager": manager}])
            cur = {t: cur_scores[t][0]["score"] for t in targets}
            rule = {t: number_rule(float(e[i]), t, ledger, manager) for i, t in enumerate(targets)}
            per = {"현재": cur, "임베딩": {t: float(e[i]) for i, t in enumerate(targets)}, "임베딩+규칙": rule,
                   "혼합": {t: 0.5 * cur[t] + 0.5 * rule[t] for t in targets}}
            for m, sc in per.items():
                st = stats[m]
                if is_pair:
                    st["pair_max"] = max(st["pair_max"], sc[target])
                    continue
                ranked = sorted(targets, key=lambda t: -sc[t])
                st["n"] += 1
                st["max_2nd"] = max(st["max_2nd"], sc[ranked[1]])
                if ranked[0] == target:
                    st["ok"] += 1
                    st["min_ok"] = min(st["min_ok"], sc[target])
                    st["min_gap"] = min(st["min_gap"], sc[target] - sc[ranked[1]])
                else:
                    st["fails"].append(f"[{set_name}] {ledger} → {ranked[0]} ({sc[ranked[0]]:.2f}, 정답 {sc[target]:.2f})")
                if detail and not is_pair:
                    print(f"  {m:<6} [{set_name}] {'O' if ranked[0] == target else 'X'} {sc[target]:.2f} / 2순위 {sc[ranked[1]]:.2f}  {ledger}")
    elapsed = time.time() - t0
    print(f"\n모델: {emb.model_path}  (펀드명 {n_embedded}개 임베딩, {elapsed:.1f}초)")
    print(f"{'방식':<10}{'1순위 정답':>10}{'정답 최저':>10}{'2순위 최고':>11}{'최소 격차':>10}{'번호만 다른 펀드 최고':>16}")
    for m in methods:
        s = stats[m]
        print(f"{m:<10}{s['ok']:>6}/{s['n']:<4}{s['min_ok']:>10.2f}{s['max_2nd']:>11.2f}{s['min_gap']:>10.2f}{s['pair_max']:>16.2f}")
    for m in methods:
        for f in stats[m]["fails"]:
            print(f"  오답({m}): {f}")
    return stats


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--model", required=True, help="ONNX 모델 폴더 (tokenizer.json + model.onnx)")
    p.add_argument("--prefix", default="", help='E5 계열은 "query: " 를 붙여야 성능이 납니다')
    p.add_argument("--cases", help="실제 종목 케이스 JSON (저장소 밖)")
    p.add_argument("--detail", action="store_true", help="케이스별 점수 출력")
    a = p.parse_args()
    evaluate(OnnxEmbedder(a.model, a.prefix), load_sets(a.cases), a.detail)


if __name__ == "__main__":
    main()
