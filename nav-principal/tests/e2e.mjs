// 실제 장부로 끝까지 돌려 보는 점검 스크립트 (장부 파일은 저장소에 넣지 않는다)
// 사용법: npm run build && TARGET=대상.xlsx LEDGERS=장부폴더 node tests/e2e.mjs
import { chromium } from "playwright";
import { readdirSync, mkdirSync, readFileSync } from "node:fs";
import { basename } from "node:path";
import { join, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const page_ = join(here, "..", "dist", "원본액자동입력_오프라인.html");
const { TARGET, LEDGERS } = process.env;
if (!TARGET || !LEDGERS) { console.error("TARGET, LEDGERS 환경변수를 지정하세요"); process.exit(2); }
const outDir = join(here, "output"); mkdirSync(outDir, { recursive: true });

// 경로에 한글이 있어도 되도록 파일 내용을 직접 넘긴다
const payload = (f) => ({ name: basename(f), mimeType: "application/octet-stream", buffer: readFileSync(f) });
const browser = await chromium.launch({ executablePath: process.env.CHROMIUM || undefined });
const page = await browser.newPage({ acceptDownloads: true });
const errors = [];
page.on("pageerror", (e) => errors.push(e.message));
page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
// PAGE_URL을 주면 배포된 주소(또는 로컬 서버)로 검사한다
await page.goto(process.env.PAGE_URL || pathToFileURL(page_).href);
await page.setInputFiles("#targetFile", payload(TARGET));
await page.waitForFunction(() => document.querySelector("#targetInfo").textContent.includes("행"));
await page.setInputFiles("#ledgerFiles", readdirSync(LEDGERS).map((f) => payload(join(LEDGERS, f))));
await page.waitForFunction(() => /읽음 <b>[1-9]/.test(document.querySelector("#ledgerInfo").innerHTML), null, { timeout: 60000 });
if (process.env.BASE_DATE) {
  await page.fill("#baseDate", process.env.BASE_DATE);
  await page.dispatchEvent("#baseDate", "change");
}
await page.$eval("#ledgerPanel", (d) => (d.open = true));
const ledgers = await page.$$eval("#ledgerBody tr", (trs) => trs.map((tr) => [...tr.children].map((td) => td.innerText.replace(/\s+/g, " ").trim()).join(" | ")));
console.log("== 장부별 추출값\n" + ledgers.join("\n"));
const rows = await page.$$eval("#tbody tr", (trs) => trs.map((tr) => [...tr.children].map((td) => (td.querySelector("select")?.selectedOptions[0]?.text ?? td.innerText).replace(/\s+/g, " ").trim()).join(" | ")));
console.log("== 매칭된 행\n" + rows.filter((r) => !/미매칭|제외/.test(r)).join("\n"));
console.log("요약:", await page.$eval("#summary", (e) => e.innerText.replace(/\s+/g, " ")));
// 검토 행까지 모두 체크하고 저장
await page.evaluate(() => { state.results.forEach((r) => { const l = state.ledgers.find((x) => x.id === r.ledgerId); if (l && l.value != null && r.status !== "skip") r.include = true; }); render(); });
const [dl] = await Promise.all([page.waitForEvent("download"), page.click("#btnSave")]);
const out = join(outDir, dl.suggestedFilename());
await dl.saveAs(out);
console.log("저장:", out);
console.log("오류:", errors.length ? errors : "없음");
await browser.close();
