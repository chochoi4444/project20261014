"""핵심 로직을 파이썬 파일 하나로 묶는다 (내부망에 텍스트로 옮기기 위함).

    python tools/bundle_core.py [출력경로]

nav_total_assets 패키지의 engine(매칭) · ledger(장부 읽기·총자산) · workbook(대상 엑셀) ·
pipeline(판정·저장) · cli(명령줄)를 순서대로 이어 붙인다. GUI·임베딩은 빼고, 모듈 사이 import는 지운다.
필요한 것: Python 3.9+, openpyxl (.xls 장부를 읽으려면 xlrd)
"""
from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

PKG = Path(__file__).resolve().parents[1] / "nav_total_assets"
MODULES = ["engine", "ledger", "workbook", "pipeline", "cli"]
# 모듈마다 같은 이름의 작은 도우미가 있으면 하나로 합칠 때 덮어쓰이므로 이름을 바꾼다
RENAME = {"cli": {"_fmt": "_fmt_cli"}}

HEADER = '''"""총자산 자동입력 — 핵심 로직 (파이썬 단일 파일)

기준가격대장(신탁재산명세부·총계정원장 등 엑셀)에서 총자산을 찾아,
대상 엑셀의 '총자산' 칸만 채운 사본을 저장한다. 원본은 덮어쓰지 않는다.

필요한 것
    Python 3.9 이상, openpyxl   (pip install openpyxl)
    .xls 장부가 있으면 xlrd      (pip install xlrd)

사용법
    python 총자산자동입력_핵심로직.py 대상.xlsx 장부폴더 [옵션]
    python 총자산자동입력_핵심로직.py --help

핵심 규칙
    총자산   = 명세부 '자산'(자산합계·자산총계) 계정의 당일잔액(차변)
               없으면 신탁재산총자산 계정 → 순자산총액+부채(검토) 순
    장부 읽기 = '당일잔액/금일잔액' 머리행을 찾고 각 값은 가장 가까운 머리글 열로 배정
    여러 탭   = 머리글+자산·순자산 계정이 있는 탭만 원장, 펀드명이 2개 이상이면 펀드마다 장부
    결산      = 이름이 같은 결산전·결산후 파일이 모두 있으면 결산후만 사용
    매칭      = 정규화·공통어 제거 후 글자 1~3묶음 TF-IDF 코사인 60% + 글자순서(LCS) 40%
               사명(운용사)이 다르면 다른 펀드, 호수·종류가 다르면 다른 펀드, 점수 높은 쌍부터 1:1
    상태      = 자동(0.85↑·2순위와 0.08↑ 차이, 또는 0.75↑·0.4↑ 차이 + 기준일 일치) / 검토 / 수동 / 미매칭 / 제외

이 파일은 저장소의 python/tools/bundle_core.py가 nav_total_assets 패키지에서 자동으로 만든 것이다.
웹판(오프라인 HTML)·엑셀 매크로판과 같은 규칙·같은 결과를 낸다.
"""
from __future__ import annotations
'''


def module_body(name: str) -> tuple[list[str], str]:
    src = (PKG / f"{name}.py").read_text(encoding="utf-8")
    tree = ast.parse(src)
    doc = ast.get_docstring(tree) or ""
    lines = src.splitlines()
    # 맨 앞 모듈 설명 문자열은 주석으로
    if tree.body and isinstance(tree.body[0], ast.Expr) and isinstance(getattr(tree.body[0], "value", None), ast.Constant):
        lines = lines[tree.body[0].end_lineno:]
    imports, body = [], []
    for ln in lines:
        if ln.startswith("from __future__"):
            continue
        if re.match(r"from \.\w* import ", ln):
            continue
        if re.match(r"(import \w|from \w[\w.]* import )", ln):
            imports.append(ln)
            continue
        body.append(ln)
    text = "\n".join(body).strip("\n")
    for old, new in RENAME.get(name, {}).items():
        text = re.sub(rf"\b{old}\b", new, text)
    # 단일 파일로 실행할 때의 명령 이름
    text = text.replace('"python -m nav_total_assets"', "Path(sys.argv[0]).name")
    head = "\n".join(f"# {x}" if x else "#" for x in doc.splitlines())
    bar = "# " + "=" * 70
    return imports, f"\n\n{bar}\n# {name}.py\n{head}\n{bar}\n\n{text}\n"


def top_names(code: str) -> list[str]:
    out = []
    for n in ast.parse(code).body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            out.append(n.name)
        elif isinstance(n, ast.Assign):
            out += [t.id for t in n.targets if isinstance(t, ast.Name)]
    return out


def build() -> str:
    imports, parts, seen = [], [], {}
    for m in MODULES:
        imp, body = module_body(m)
        for x in imp:
            if x not in imports:
                imports.append(x)
        for n in top_names(body):
            if n in seen:
                raise SystemExit(f"이름 충돌: {n} ({seen[n]}.py, {m}.py) — RENAME에 추가하세요")
            seen[n] = m
        parts.append(body)
    imports.sort(key=lambda x: (x.startswith("from"), x))
    tail = '\n\nif __name__ == "__main__":\n    sys.exit(main())\n'
    return HEADER + "\n" + "\n".join(imports) + "\n" + "".join(parts) + tail


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("총자산자동입력_핵심로직.py")
    code = build()
    compile(code, str(out), "exec")
    out.write_text(code, encoding="utf-8")
    print(f"만들었습니다: {out} ({len(code.splitlines())}줄)")
