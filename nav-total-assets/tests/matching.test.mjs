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
const ctx = vm.createContext({});
vm.runInContext(code + "\n;globalThis.E = { scoreName, matchNames, matchStatus, normName, coreName, nums, cosineSim, buildIdf, AUTO_SCORE, REVIEW_SCORE };", ctx);
const E = ctx.E;
const STATUS_KO = { auto: "자동", review: "검토", nomatch: "미매칭" };

// 대상 엑셀 종목명 목록 (가상). 이름이 비슷하고 번호만 다른 펀드를 일부러 섞었다.
const TARGETS = [
  "가온대체투자전문투자형사모부동산투자신탁101호",
  "가온대체투자전문투자형사모부동산투자신탁102호 1종",
  "가온대체투자전문투자형사모부동산투자신탁102호 2종",
  "가온대체투자전문투자형사모부동산투자신탁 67호",
  "가온대체투자전문투자형사모부동산투자신탁67-1호",
  "가온대체투자전문투자형사모부동산투자신탁89호1종",
  "가온대체투자에너지인프라전문투자형사모특별자산투자신탁3-1호",
  "가온대체투자에너지인프라전문투자형사모특별자산투자신탁3-3호[재간접]",
  "가온대체투자블랙락전문투자형사모73호",
  "가온1호전문투자사모부동산유한회사",
  "가온2호전문투자사모부동산유한회사",
  "누리인베스트MACKUSDEBT사모부동산4호",
  "다온북미AG일반사모특별자산신탁제1호(대출채권-재간접형)(전문)",
  "채권시장안정사모증권투자신탁 [채권-재간접형]",
  "보람북미크레딧전문투자사모특별자산8",
  "보람북미크레딧전문투자형사모특별자산투자신탁26호(재간접형)",
  "새봄 USDL Ⅲ LEVERED NOTE 일반사모투자신탁 제1호",
  "마루더드림중단기크레딧증권투자신탁[채권]CLASSC-F",
  "마루다이렉트초단기채증권투자신탁[채권]C-F",
  "솔빛대체일반사모특별자산투자신탁82호",
  "솔빛대체일반사모특별자산투자신탁84호",
  "솔빛대체일반사모특별자산투자신탁86호",
  "하람유럽인프라전문투자형사모특별자산투자신탁1호(대출채권-재간접형)",
  "하람모두성장인프라일반사모특별자산투자신탁1호",
  "이음해외간접사모부동산신탁21호",
  "한울파트너스전문투자형사모부동산투자신탁유럽제1호",
  "여울레포일반사모증권투자신탁5[채권]",
  "여울레포일반사모증권투자신탁6[채권]",
  "여울레포일반사모증권투자신탁7[채권]",
  "여울AIM재간접전문투자형사모투자신탁2호",
  "여울AIM재간접전문투자형사모투자신탁4호",
  "예시법인용MMF 1(국공채)C",
];

// [대상 종목명, 장부에 적힌 표기, 장부 머리글 운용사명, 허용 상태, 달라진 방식]
const VARIANTS = [
  [TARGETS[0], "가온대체투자일반사모부동산투자신탁101호", "", ["auto"], "법 개정 명칭(전문투자형사모→일반사모)"],
  [TARGETS[11], "누리인베스트먼트 Mack US Debt 일반사모부동산투자신탁 제4호", "", ["auto"], "운용사 정식명·영문 대소문자·띄어쓰기"],
  [TARGETS[12], "다온북미AG일반사모특별자산투자신탁1호", "", ["auto"], "괄호 부가설명 생략"],
  [TARGETS[13], "채권시장안정일반사모증권투자신탁(채권-재간접형)", "", ["auto"], "괄호 종류·사모 표기"],
  [TARGETS[14], "보람북미크레딧일반사모특별자산투자신탁8호", "", ["auto"], "'호' 생략 vs 표기"],
  [TARGETS[8], "가온대체투자 BlackRock 일반사모투자신탁 73호", "", ["auto"], "한글↔영문(블랙락)"],
  [TARGETS[16], "새봄 USDL III Levered Note 일반사모투자신탁 제1호", "", ["auto"], "로마숫자·대소문자"],
  [TARGETS[17], "마루더드림중단기크레딧증권투자신탁[채권]C-F", "", ["auto"], "클래스 표기"],
  [TARGETS[21], "솔빛대체투자일반사모특별자산투자신탁 제86호", "", ["auto"], "운용사명 일부·'제' 표기"],
  [TARGETS[22], "하람 유럽 인프라 일반사모 특별자산 투자신탁 1호", "", ["auto"], "띄어쓰기·괄호 생략"],
  [TARGETS[25], "한울해외부동산투자신탁1호(전문)", "한울파트너스자산운용", ["auto", "review"], "운용사 약칭·단어 차이(머리글 운용사명으로 보강)"],
  [TARGETS[24], "이음글로벌일반사모부동산투자신탁21호", "", ["review"], "핵심 단어가 바뀜(해외간접→글로벌)"],
  [TARGETS[26], "여울레포일반사모증권투자신탁5(채권)", "", ["auto"], "괄호 모양"],
  [TARGETS[5], "가온대체투자일반사모부동산투자신탁89호(1종)", "", ["auto"], "종류 괄호"],
];

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

// [대상 종목명, 같은 이름에 번호만 다른 장부]
const NUMBER_PAIRS = [
  [TARGETS[26], "여울레포일반사모증권투자신탁6(채권)", "5 vs 6"],
  [TARGETS[9], "가온2호일반사모부동산유한회사", "1호 vs 2호"],
  [TARGETS[29], "여울AIM재간접일반사모투자신탁4호", "2호 vs 4호"],
  [TARGETS[19], "솔빛대체일반사모특별자산투자신탁86호", "82호 vs 86호"],
  [TARGETS[5], "가온대체투자일반사모부동산투자신탁89호(2종)", "89호1종 vs 89호2종"],
  [TARGETS[3], "가온대체투자일반사모부동산투자신탁67-1호", "67호 vs 67-1호"],
  [TARGETS[6], "가온대체투자에너지인프라일반사모특별자산투자신탁3-3호", "3-1호 vs 3-3호"],
  [TARGETS[1], "가온대체투자일반사모부동산투자신탁102호 2종", "102호 1종 vs 2종"],
];

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
  const { assign } = E.matchNames(TARGETS, [ledger(0, "가온대체투자일반사모부동산투자신탁89호(2종)")]);
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

test("저장된 수동 매핑이 점수보다 우선한다", () => {
  const L = [ledger(0, "전혀 다른 이름의 장부"), ledger(1, "가온대체투자일반사모부동산투자신탁101호")];
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
