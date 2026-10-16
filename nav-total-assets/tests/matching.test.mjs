// 펀드명 매칭 회귀 테스트: npm test
// 이 저장소는 공개 저장소라 실제 보유 펀드명 대신, 같은 방식으로 표기가 달라지는 가상의 펀드명을 쓴다.
// 실제 종목으로도 돌려 보려면 저장소 밖에 JSON을 두고 MATCH_CASES=경로 로 지정한다 (형식은 맨 아래 참고).
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, "..", "src", "app.html"), "utf8");
const START = "/* ==== 매칭 엔진 시작", END = "/* ==== 매칭 엔진 끝 ==== */";
assert.ok(html.includes(START) && html.includes(END), "app.html에서 매칭 엔진 구역 표시를 찾지 못했습니다");
const code = html.slice(html.indexOf(START), html.indexOf(END));
const cleanFn = html.slice(html.indexOf("function cleanFileName"), html.indexOf("/* ---------- 장부 읽기"));
const ctx = vm.createContext({});
vm.runInContext(code + "\n" + cleanFn + "\n;globalThis.E = { cleanFileName, scoreName, matchNames, matchStatus, normName, coreName, nums, cosineSim, buildIdf, AUTO_SCORE, REVIEW_SCORE };", ctx);
const E = ctx.E;
const STATUS_KO = { auto: "자동", review: "검토", nomatch: "미매칭" };

// 가상 펀드명 케이스 (파이썬 테스트와 공유)
const FIX = JSON.parse(readFileSync(join(here, "fixtures", "matching-cases.json"), "utf8"));
const TARGETS = FIX.targets;
const VARIANTS = FIX.variants.map((v) => [v.target, v.ledger, v.manager, v.allowed, v.how]);

const ledger = (id, name, manager = "") => ({ id, nameCands: [name], name, manager });

test("표기가 달라도 맞는 종목을 1순위로 고른다", async (t) => {
  for (const [target, variant, manager, allowed, how] of VARIANTS) {
    await t.test(how, () => {
      const L = ledger(0, variant, manager);
      // 이 장부 하나만 있을 때의 점수(TF-IDF 포함)·배정·상태
      const { scores, assign } = E.matchNames(TARGETS, [L]);
      const ranked = TARGETS.map((n) => ({ n, score: scores.get(n)[0].score })).sort((a, b) => b.score - a.score);
      assert.equal(ranked[0].n, target, `1순위가 ${ranked[0].n} (${ranked[0].score.toFixed(2)})`);
      assert.equal(assign.get(target), 0, "1:1 배정에서 빠졌습니다");
      const st = E.matchStatus(target, 0, scores);
      assert.ok(allowed.includes(st), `상태 ${STATUS_KO[st]} (점수 ${ranked[0].score.toFixed(2)}), 기대 ${allowed.map((x) => STATUS_KO[x]).join("/")}`);
    });
  }
});

const NUMBER_PAIRS = FIX.number_pairs.map((p) => [p.target, p.ledger, p.how]);

test("이름이 같아도 호수·종류가 다르면 후보가 되지 않는다", async (t) => {
  for (const [target, other, how] of NUMBER_PAIRS) {
    await t.test(how, () => {
      const s = E.matchNames(TARGETS, [ledger(0, other)]).scores.get(target)[0].score;
      assert.ok(s < E.REVIEW_SCORE, `점수 ${s.toFixed(2)} ≥ ${E.REVIEW_SCORE}`);
    });
  }
});

test("맞는 장부가 빠져 있으면 번호만 다른 펀드에 붙이지 않는다", () => {
  // 89호1종 장부는 없고 89호2종 장부만 있는 경우: 대상 목록에 89호2종이 없으므로 어디에도 붙으면 안 된다
  const { assign } = E.matchNames(TARGETS, [ledger(0, FIX.number_pairs.find((p) => p.how.startsWith("89호")).ledger)]);
  for (const [n, id] of assign) assert.notEqual(id, 0, `${n}에 잘못 붙었습니다`);
});

test("여러 장부를 한꺼번에 올리면 각자 맞는 종목에 1:1로 붙는다", () => {
  const all = VARIANTS.map(([, v, m], i) => ledger(i, v, m)).concat(
    NUMBER_PAIRS.map(([, o], i) => ledger(100 + i, o)).filter((l) => !TARGETS.some((t) => E.normName(t) === E.normName(l.name))),
  );
  // 번호만 다른 장부 중 대상 목록에 실제로 있는 펀드(예: 6, 2호, 4호, 86호, 67-1호, 3-3호, 102호 2종)는 그 펀드에 붙어야 한다
  const { assign } = E.matchNames(TARGETS, all);
  VARIANTS.forEach(([target], i) => assert.equal(assign.get(target), i, `${target} ← ${assign.get(target)}`));
  const used = [...assign.values()].filter((v) => v != null);
  assert.equal(new Set(used).size, used.length, "한 장부가 두 종목에 붙었습니다");
});

test("코사인 유사도: 같은 이름 1, 겹치는 글자 없으면 0, 흔한 글자는 가중치가 낮다", () => {
  assert.equal(E.cosineSim("가온블랙락73", "가온블랙락73").toFixed(6), "1.000000");
  assert.equal(E.cosineSim("가온블랙락", "누리USDL"), 0);
  // '부동산'처럼 여러 이름에 흔한 글자로 겹친 쌍은, 고유한 글자('가온')로 겹친 쌍보다 TF-IDF에서 점수가 더 많이 깎인다
  const idf = E.buildIdf(["가온부동산", "누리부동산", "다온부동산", "가온인프라"]);
  const ratio = (a, b) => E.cosineSim(a, b, idf) / E.cosineSim(a, b);
  assert.ok(ratio("가온부동산", "누리부동산") < ratio("가온부동산", "가온인프라"));
});

test("파일명의 날짜·펀드코드·사본 표시를 지우고 호수만 남긴다", async (t) => {
  const { cases } = JSON.parse(readFileSync(join(here, "fixtures", "filename-cases.json"), "utf8"));
  for (const c of cases) {
    await t.test(c.file, () => {
      const title = E.cleanFileName(c.file);
      assert.equal(title, c.title);
      assert.deepEqual([...E.nums(title)], c.nums);
    });
  }
});

test("파일명에 날짜·펀드코드가 있어도 맞는 종목에 붙는다", () => {
  const L = (name) => ({ id: 0, nameCands: [E.cleanFileName(name)], name: null, fileTitle: E.cleanFileName(name), manager: "" });
  for (const [file, target] of [
    ["기준가격대장_2026.09.30_보람북미크레딧일반사모특별자산투자신탁8호.xls", TARGETS[14]],
    ["2026-09-30 가온대체투자일반사모부동산투자신탁101호.xlsx", TARGETS[0]],
    ["18021_가온대체투자일반사모부동산투자신탁101호.xls", TARGETS[0]],
  ]) {
    const { assign } = E.matchNames(TARGETS, [L(file)]);
    assert.equal(assign.get(target), 0, `${file} → 배정 안 됨`);
  }
});

test("저장된 수동 매핑이 점수보다 우선한다", () => {
  const L = [ledger(0, "전혀 다른 이름의 장부"), ledger(1, FIX.variants[0].ledger)];
  const { assign } = E.matchNames(TARGETS, L, new Map([[TARGETS[0], 0]]));
  assert.equal(assign.get(TARGETS[0]), 0);
  assert.notEqual(assign.get(TARGETS[0]), 1);
});

// 실제 종목으로 검사 (선택): {"targets": ["종목명", ...], "cases": [["대상 종목명", "장부 표기", "운용사명(선택)", ["auto","review"](선택)], ...]}
const real = process.env.MATCH_CASES;
test("실제 종목 케이스 (MATCH_CASES)", { skip: !real || !existsSync(real) ? "MATCH_CASES 미지정" : false }, () => {
  const { targets, cases } = JSON.parse(readFileSync(real, "utf8"));
  for (const [target, variant, manager = "", allowed = ["auto", "review"]] of cases) {
    const L = ledger(0, variant, manager);
    const { scores } = E.matchNames(targets, [L]);
    const ranked = targets.map((n) => ({ n, score: scores.get(n)[0].score })).sort((a, b) => b.score - a.score);
    assert.equal(ranked[0].n, target, `${variant} → ${ranked[0].n}`);
    assert.ok(allowed.includes(E.matchStatus(target, 0, scores)), `${variant}: ${E.matchStatus(target, 0, scores)}`);
  }
});
