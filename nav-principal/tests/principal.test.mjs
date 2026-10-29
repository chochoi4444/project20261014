// 원본액 산출 회귀 테스트: npm test
// app.html의 장부 읽기 구역(계정 찾기·원본액 산출)만 떼어 가상의 시산표 행으로 검사한다.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, "..", "src", "app.html"), "utf8");
const cut = (a, b) => html.slice(html.indexOf(a), html.indexOf(b));
const code = [
  cut("const fmt =", "const fmtNav"),
  cut("/* ==== 매칭 엔진 시작", "/* ==== 매칭 엔진 끝 ==== */"),
  cut("/* ---------- 숫자·날짜 ---------- */", "/* ---------- 대상 엑셀 ---------- */"),
].join("\n");
const ctx = vm.createContext({});
vm.runInContext(code + "\n;globalThis.E = { extractFromRows, computePrincipal };", ctx);
const E = ctx.E;

// 시산표 한 장 (전일차변 | 전일대변 | 당일차변 | 당일대변): [계정명, 전일차, 전일대, 당일차, 당일대]
const sheet = (lines) => [
  [{ c: 0, v: "계정과목" }, { c: 1, v: "전일차변잔액" }, { c: 2, v: "전일대변잔액" }, { c: 3, v: "당일차변잔액" }, { c: 4, v: "당일대변잔액" }],
  ...lines.map(([label, ...vals]) => [{ c: 0, v: label }, ...vals.map((v, k) => (v == null ? null : { c: k + 1, v })).filter(Boolean)]),
];
const finish = (ex) => { const L = { fields: ex.fields, hasPrev: ex.hasPrev }; E.computePrincipal(L); return L; };
const run = (lines) => finish(E.extractFromRows(sheet(lines)));

test("원본액 계정의 전일 대변잔액을 읽는다", () => {
  const L = run([["자산", 12_000, null, 12_500, null], ["원본액", null, 10_000_000_000, null, 10_300_000_000], ["신탁재산순자산총액", null, 1, null, 2]]);
  assert.equal(L.value, 10_000_000_000);
  assert.match(L.method, /원본액 전일 대변잔액/);
  assert.equal(L.check.s, "mute", "잔존좌수가 없으면 대조 안 함");
  assert.match(L.check.t, /300,000,000/);
});

test("잔존좌수가 없으면 대조하지 않고 그대로 넣는다", () => {
  const L = run([["원본액", null, 5_000, null, 5_000], ["자산", 1, null, 1, null]]);
  assert.equal(L.value, 5_000);
  assert.equal(L.units, null);
  assert.equal(L.check.s, "mute");
});

test("전일잔액 병합 머리글 + 아랫줄 차변/대변 양식", () => {
  const rows = [
    [{ c: 0, v: "계정과목" }, { c: 1, v: "전일잔액" }, { c: 3, v: "당일증감" }, { c: 5, v: "당일잔액" }],
    [{ c: 1, v: "차변" }, { c: 2, v: "대변" }, { c: 3, v: "차변" }, { c: 4, v: "대변" }, { c: 5, v: "차변" }, { c: 6, v: "대변" }],
    [{ c: 0, v: "자산" }, { c: 1, v: 900 }, { c: 5, v: 950 }],
    [{ c: 0, v: "원본액" }, { c: 2, v: 7_000 }, { c: 4, v: 100 }, { c: 6, v: 7_100 }],
  ];
  const L = finish(E.extractFromRows(rows));
  assert.equal(L.value, 7_000);
  assert.equal(L.fields.principal.v, 7_100, "당일 대변");
});

test("전일 대변 칸이 비어 있으면 값을 넣지 않는다", () => {
  const L = run([["원본액", null, null, null, 8_000], ["자산", 1, null, 1, null]]);
  assert.equal(L.value, null);
  assert.equal(L.check.s, "bad");
});

test("원본액 계정이 없으면 값이 없다", () => {
  const L = run([["자산", 9, null, 9, null], ["수익증권잔존좌수", 7, null, 7, null]]);
  assert.equal(L.value, null);
  assert.match(L.check.t, /원본액 계정/);
});

test("원본이익 같은 다른 계정은 원본액으로 보지 않는다", () => {
  const L = run([["원본이익", null, 1_234, null, 1_234], ["원본액", null, 2_000, null, 2_000]]);
  assert.equal(L.value, 2_000);
});

test("투자회사는 원본액 대신 잔존좌수(발행주식수)를 쓰고 검토로 표시한다", () => {
  const ex = E.extractFromRows(sheet([["원본액", null, 327_000, null, 327_000], ["잔존좌수", null, 86_649, null, 86_649]]));
  const L = { fields: ex.fields, hasPrev: ex.hasPrev, name: "가온이머징인프라투융자회사", meta: {} };
  E.computePrincipal(L);
  assert.equal(L.value, 86_649);
  assert.equal(L.check.s, "warn");
});

test("원본액과 잔존좌수(차변 쪽)가 같으면 일치, 다르면 검토", () => {
  const same = run([["원본액", null, 5_000, null, 5_000], ["수익증권잔존좌수", 5_000, null, 5_000, null]]);
  assert.equal(same.units, 5_000);
  assert.equal(same.check.s, "ok", same.check.t);
  const diff = run([["원본액", null, 5_000, null, 5_000], ["잔존좌수", null, 4_000, null, 4_000]]);
  assert.equal(diff.value, 5_000);
  assert.equal(diff.units, 4_000);
  assert.equal(diff.check.s, "warn");
  assert.match(diff.check.t, /1,000 차이/);
});
