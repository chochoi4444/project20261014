"""웹 버전(src/app.html)과 파이썬 버전이 같은 점수를 내는지 대조한다. node가 없으면 건너뛴다."""
import json
import shutil
import subprocess
from pathlib import Path

import pytest

from nav_total_assets.engine import match_names, match_status

APP = Path(__file__).resolve().parents[2] / "src" / "app.html"

JS = r"""
const fs = require("fs"), vm = require("vm");
const html = fs.readFileSync(process.argv[1], "utf8");
const code = html.slice(html.indexOf("/* ==== 매칭 엔진 시작"), html.indexOf("/* ==== 매칭 엔진 끝 ==== */"));
const ctx = vm.createContext({});
vm.runInContext(code + ";globalThis.E = { matchNames, matchStatus };", ctx);
const { targets, ledgers } = JSON.parse(fs.readFileSync(0, "utf8"));
const { scores, assign } = ctx.E.matchNames(targets, ledgers);
const out = { scores: {}, assign: {}, status: {} };
for (const [n, list] of scores) out.scores[n] = list.map((e) => [e.id, e.score, e.why]);
for (const [n, id] of assign) { out.assign[n] = id; out.status[n] = ctx.E.matchStatus(n, id, scores); }
process.stdout.write(JSON.stringify(out));
"""


@pytest.mark.skipif(not shutil.which("node"), reason="node 없음")
def test_same_scores_as_web(cases):
    T = cases["targets"]
    ls = [{"id": i, "nameCands": [v["ledger"]], "name": v["ledger"], "manager": v["manager"]} for i, v in enumerate(cases["variants"])]
    ls += [{"id": 100 + i, "nameCands": [p["ledger"]], "name": p["ledger"], "manager": ""} for i, p in enumerate(cases["number_pairs"])]
    js = json.loads(subprocess.run(["node", "-e", JS, str(APP)], input=json.dumps({"targets": T, "ledgers": ls}),
                                   capture_output=True, text=True, check=True).stdout)
    py_ls = [{"id": l["id"], "name_cands": l["nameCands"], "name": l["name"], "manager": l["manager"]} for l in ls]
    scores, assign = match_names(T, py_ls)
    n = 0
    for name in T:
        for (jid, jscore, jwhy), e in zip(js["scores"][name], scores[name]):
            assert jid == e["id"]
            assert e["score"] == pytest.approx(jscore, abs=1e-9), f"{name} / 장부 {jid}"
            assert e["why"] == jwhy
            n += 1
    assert {k: v for k, v in assign.items()} == js["assign"]
    assert {k: match_status(k, v, scores) for k, v in assign.items()} == js["status"]
    assert n == len(T) * len(ls)
