"""웹앱(src/app.html)의 핵심 로직만 그대로 뽑아 .txt로 저장한다.

    python tools/extract_main_logic.py [출력경로]

- 화면 디자인/스타일(<style>, HTML 마크업, 화면 그리기·이벤트 코드)과 기능 설명 화면은 뺀다.
- 남기는 코드는 고치거나 줄이지 않고 원문 그대로 옮긴다 (최상위 선언 단위로 넣거나 뺀다).
- API 키·비밀번호처럼 보이는 값이 있으면 실제 값을 가린다.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src" / "app.html"

# 화면(DOM) 그리기·이벤트·저장 창 처리 등 핵심 로직이 아닌 최상위 선언
UI = {
    "downloadsP", "$", "esc", "fmtNav", "MAP_KEY", "localStorage", "saveMap", "updateMapCount",
    "STATUS", "DOTC", "scoreCell", "tabsNote", "unmatchedCell", "render", "busy", "toastTimer", "toast",
    "renderLedgerInfo", "filesFromDrop", "dropzones", "inArtifactViewer", "SAVE_BLOCKED", "saveFile",
    "showTab", "libcheck", "onTarget",
}
SECRET = re.compile(r"""(?i)\b(api[_-]?key|secret|token|password|passwd|pwd|authorization|bearer)\b(\s*[:=]\s*)(["'`])([^"'`]{6,})\3""")


def script_body(html: str) -> str:
    blocks = re.findall(r"<script>(.*?)</script>", html, re.S)  # src= 로 불러오는 라이브러리는 제외
    if len(blocks) != 1:
        raise SystemExit(f"인라인 스크립트가 {len(blocks)}개입니다 (1개여야 함)")
    return blocks[0].strip("\n")


def block_name(first: str) -> str:
    m = re.match(r"(?:async\s+)?function\s+([\w$]+)", first) or re.match(r"(?:const|let|var)\s+([\w$]+)", first)
    if m:
        return m[1]
    m = re.match(r'\$\("([^"]+)"\)\.addEventListener', first)
    if m:
        return "event:" + m[1]
    if first.startswith(("document.addEventListener", "window.addEventListener")):
        return "event:" + first.split(".")[0]
    m = re.match(r"([\w$]+)\(.*\);$", first)  # 최상위에서 바로 부르는 초기화 호출
    if m:
        return "call:" + m[1]
    if first.startswith("for (const [id, handler]"):
        return "dropzones"
    if first.startswith("try { localStorage"):
        return "localStorage"
    if first.startswith("if (!window.XLSX"):
        return "libcheck"
    return first[:40]


def split_blocks(js: str):
    """최상위(들여쓰기 없는) 문장 단위로 나눈다. 바로 위의 주석은 그 문장에 붙인다."""
    lines = js.splitlines()
    blocks, cur, pending, in_comment = [], None, [], False
    for ln in lines:
        top = ln[:1] not in ("", " ", "\t", "}", "]", ")")
        if in_comment:
            pending.append(ln)
            if "*/" in ln:
                in_comment = False
            continue
        if top and (ln.startswith("//") or ln.startswith("/*")):
            if cur is not None:
                blocks.append(cur)
                cur = None
            pending.append(ln)
            in_comment = ln.startswith("/*") and "*/" not in ln
            continue
        if top:
            if cur is not None:
                blocks.append(cur)
            cur = {"name": block_name(ln), "lines": pending + [ln]}
            pending = []
            continue
        if cur is None:
            pending.append(ln)
        else:
            cur["lines"].append(ln)
    if cur is not None:
        blocks.append(cur)
    if pending:
        blocks.append({"name": "comment", "lines": pending})
    return blocks


def is_ui(b) -> bool:
    n = b["name"]
    return n in UI or n.startswith("event:") or (n.startswith("call:") and n[5:] in UI)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist" / "총자산자동입력_핵심로직_웹앱.txt"
    js = script_body(SRC.read_text(encoding="utf-8"))
    blocks = split_blocks(js)
    # 나눈 조각을 다시 이으면 원문과 같아야 한다 (빠뜨리거나 바꾼 줄 없음)
    assert "\n".join(l for b in blocks for l in b["lines"]) == js, "블록 분할이 원문과 다릅니다"
    kept = [b for b in blocks if not is_ui(b)]
    body = "\n".join(l for b in kept for l in b["lines"]).strip("\n") + "\n"
    body = re.sub(r"\n{3,}", "\n\n", body)  # 뺀 자리에 남는 빈 줄만 정리
    body, n_secret = SECRET.subn(lambda m: m[1] + m[2] + m[3] + "***REDACTED***" + m[3], body)
    try:
        rev = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, capture_output=True, text=True).stdout.strip()
    except OSError:
        rev = ""
    head = (f"// 출처: nav-total-assets/src/app.html{f' (커밋 {rev})' if rev else ''} — <script> 원문에서 핵심 로직만 그대로 추출\n"
            f"// 제외: <style>·HTML 마크업·기능 설명 화면, 화면 그리기/이벤트 코드 {sum(map(is_ui, blocks))}개 블록"
            f" ({', '.join(b['name'] for b in blocks if is_ui(b))})\n"
            f"// 외부 라이브러리(SheetJS xlsx 0.18.5, JSZip 3.10.1)는 포함하지 않음 · 인증정보 가림 {n_secret}건\n\n")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(head + body, encoding="utf-8")
    print(f"만들었습니다: {out} ({len((head + body).splitlines())}줄, 블록 {len(kept)}/{len(blocks)}개 유지)")


if __name__ == "__main__":
    main()
