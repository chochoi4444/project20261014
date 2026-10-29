// 오프라인판 생성: 외부 라이브러리 2개를 HTML 하나에 넣어 인터넷 없이 열 수 있게 만든다.
// 사용법: npm install && npm run build  →  dist/원본액자동입력_오프라인.html
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const lib = (p) => readFileSync(join(here, "node_modules", p), "utf8").replace(/<\/script/gi, "<\\/script");
const libs = {
  "xlsx/0.18.5/xlsx.full.min.js": lib("xlsx/dist/xlsx.full.min.js"),
  "jszip/3.10.1/jszip.min.js": lib("jszip/dist/jszip.min.js"),
};

let html = readFileSync(join(here, "src/app.html"), "utf8");
// 웹 폰트도 빼서 외부 접속을 0건으로 만든다 (시스템 글꼴로 대체)
html = html.replace(/<link\b[^>]*fonts\.(googleapis|gstatic)\.com[^>]*>\n?/g, "");
for (const [path, code] of Object.entries(libs)) {
  const tag = `<script src="https://cdnjs.cloudflare.com/ajax/libs/${path}"></script>`;
  if (!html.includes(tag)) throw new Error("CDN 태그를 찾지 못함: " + path);
  html = html.replace(tag, () => `<script>${code}</script>`);
}
if (/https?:\/\/(?!www\.w3\.org)[^\s"'<>]*\.(js|css)\b/.test(html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, ""))) {
  throw new Error("외부 리소스 참조가 남아 있습니다");
}
const out = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>[hidden]{display:none!important}body{margin:0}</style></head><body>\n${html}\n</body></html>\n`;
mkdirSync(join(here, "dist"), { recursive: true });
const file = join(here, "dist", "원본액자동입력_오프라인.html");
writeFileSync(file, out);
// Vercel 등 정적 호스팅용: 같은 내용을 index.html로도 둔다
writeFileSync(join(here, "dist", "index.html"), out);
writeFileSync(join(here, "dist", "robots.txt"), "User-agent: *\nDisallow: /\n");
console.log(`만들었습니다: ${file} (${(out.length / 1024 / 1024).toFixed(1)} MB) + dist/index.html`);
