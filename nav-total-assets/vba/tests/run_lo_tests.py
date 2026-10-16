"""엑셀 매크로(TotalAssets.bas)의 계산부를 LibreOffice(VBA 호환 모드)로 돌려 파이썬판과 대조한다.

    python run_lo_tests.py                       # 매칭 엔진 대조 + 가상 장부 끝까지
    TARGET=대상.xlsx LEDGERS=장부폴더 BASE=2026-09-30 python run_lo_tests.py   # 실제 파일 추가

LibreOffice(soffice)가 필요하다. 엑셀의 파일 열기·저장 부분은 여기서 돌지 않는다.
"""
import html, json, os, subprocess, sys, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT / "python"))
sys.path.insert(0, str(ROOT / "python" / "tests"))
from nav_total_assets.engine import match_names, match_status  # noqa: E402
from nav_total_assets.pipeline import run  # noqa: E402

PROFILE = Path(tempfile.gettempdir()) / "nav-lo-profile"


def install_modules():
    if not (PROFILE / "user/basic/Standard").exists():
        subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{PROFILE}", "--terminate_after_init"],
                       capture_output=True, timeout=180)
    lib = PROFILE / "user/basic/Standard"
    src = (HERE.parent / "TotalAssets.bas").read_text(encoding="utf-8")
    harness = (HERE / "lo_harness.bas").read_text(encoding="utf-8")
    harness = "\n".join(l for l in harness.split("\n") if not l.startswith("Option "))
    # LibreOffice는 다른 모듈에서 들어온 호출이 이어질 때 모듈 변수를 잃는다 (엑셀은 문제없음).
    # 그래서 점검용 함수를 매크로 모듈 끝에 붙여 한 모듈 안에서 돌린다.
    mods = {"Harness": "Option VBASupport 1\n" + src + "\n" + harness}
    for name, code in mods.items():
        (lib / f"{name}.xba").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            f'<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{name}" script:language="StarBasic">'
            + html.escape(code, quote=False) + "</script:module>", encoding="utf-8")
    els = "\n".join(f' <library:element library:name="{n}"/>' for n in mods)
    (lib / "script.xlb").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
        '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
        + els + "\n</library:library>\n", encoding="utf-8")


_CTX = None


def _connect():
    """soffice를 소켓으로 띄우고 UNO로 붙는다 (Basic 오류가 대화상자 대신 예외로 온다)"""
    global _CTX
    if _CTX:
        return _CTX
    import time
    import uno
    subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", f"-env:UserInstallation=file://{PROFILE}",
                      "--accept=socket,host=127.0.0.1,port=2083;urp;"], stdout=subprocess.DEVNULL, stderr=open(Path(tempfile.gettempdir()) / "nav-lo-stderr.log", "w"))
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            _CTX = resolver.resolve("uno:socket,host=127.0.0.1,port=2083;urp;StarOffice.ComponentContext")
            return _CTX
        except Exception:
            time.sleep(0.5)
    raise SystemExit("LibreOffice에 연결하지 못했습니다")


def run_macro(name, *args):
    ctx = _connect()
    smgr = ctx.ServiceManager
    mspf = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
    sp = mspf.createScriptProvider("")
    script = sp.getScript(f"vnd.sun.star.script:Standard.Harness.{name}?language=Basic&location=application")
    result, _, _ = script.invoke(tuple(args), (), ())
    return result


def shutdown():
    if _CTX:
        try:
            _CTX.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", _CTX).terminate()
        except Exception:
            pass


def parity():
    fx = json.loads((ROOT / "tests/fixtures/matching-cases.json").read_text(encoding="utf-8"))
    T = fx["targets"]
    L = [(v["ledger"], v["manager"]) for v in fx["variants"]] + [(p["ledger"], "") for p in fx["number_pairs"]]
    txt = run_macro("Parity", "".join(f"T\t{t}\n" for t in T) + "".join(f"L\t{n}\t{m}\n" for n, m in L))
    vb_s, vb_a = {}, {}
    for ln in txt.splitlines():
        p = ln.split("\t")
        if p[0] == "S":
            vb_s[(int(p[1]), int(p[2]))] = (float(p[3]), p[4])
        elif p[0] == "A":
            vb_a[int(p[1])] = (int(p[2]), p[3])
    scores, assign = match_names(T, [{"id": i, "name_cands": [n], "name": n, "manager": m} for i, (n, m) in enumerate(L)])
    bad = 0
    for ti, t in enumerate(T):
        for e in scores[t]:
            vs, vw = vb_s[(ti, e["id"])]
            if abs(vs - e["score"]) > 1e-9 or vw != e["why"]:
                bad += 1
                if bad <= 5:
                    print(f"  다름: {t} / {L[e['id']][0]}: 매크로 {vs:.6f} '{vw}' vs 파이썬 {e['score']:.6f} '{e['why']}'")
    py_a = {ti: (assign[t], match_status(t, assign[t], scores)) for ti, t in enumerate(T) if assign.get(t) is not None}
    if py_a != vb_a:
        bad += 1
        print("  배정·상태 다름:", py_a, vb_a)
    print(f"매칭 엔진 대조: {len(T)}×{len(L)} = {len(T) * len(L)}쌍, 다른 곳 {bad}")
    return bad == 0


def filenames():
    data = json.loads((ROOT / "tests/fixtures/filename-cases.json").read_text(encoding="utf-8"))
    txt = run_macro("CleanNames", "".join(c["file"] + "\n" for c in data["cases"]))
    rows = txt.split("\n")
    bad = 0
    for c, row in zip(data["cases"], rows):
        title, key = (row.split("\t") + [""])[:2]
        if title != c["title"] or key != "|".join(c["nums"]):
            bad += 1
            print(f"  다름: {c['file']}: 매크로 '{title}' [{key}] vs 기대 '{c['title']}' {c['nums']}")
    print(f"파일명 정리 대조: {len(data['cases'])}건, 다른 곳 {bad}")
    return bad == 0


def full(target, ledgers, base):
    txt = run_macro("Full", str(target), str(ledgers), base or "")
    _, lds, results, base_used = run(target, [ledgers], base or None)
    vb_rows = [ln.split("\t") for ln in txt.splitlines() if ln.startswith("R\t")]
    vb_led = [ln.split("\t") for ln in txt.splitlines() if ln.startswith("L\t")]
    bad = 0
    py_led = sorted(lds, key=lambda l: l.file_name)
    for v, l in zip(vb_led, py_led):
        pv = "" if l.value is None else str(round(l.value))
        exp = [l.file_name, l.name or "", l.file_title, l.manager, l.code, l.date or "", pv, l.method, l.check[1], l.sheet or ""]
        if v[1:] != exp:
            bad += 1
            print("  장부 다름:\n    매크로", v[1:], "\n    파이썬", exp)
    for v, r in zip(vb_rows, results):
        exp = [str(r.excel_row), r.status, "Y" if r.include else "", f"{r.score:.6f}", r.ledger.file_name if r.ledger else "", r.why]
        if v[1:7] != exp:
            bad += 1
            if bad <= 8:
                print("  행 다름:\n    매크로", v[1:7], "\n    파이썬", exp)
    if len(vb_rows) != len(results) or len(vb_led) != len(lds):
        bad += 1
        print("  개수 다름", len(vb_rows), len(results), len(vb_led), len(lds))
    auto = [v for v in vb_rows if v[2] in ("auto", "review", "manual")]
    print(f"끝까지 대조 ({Path(ledgers).name}, 기준일 {base or '머리글'}): 장부 {len(vb_led)}개, 행 {len(vb_rows)}개, 다른 곳 {bad}")
    for v in auto:
        print(f"    행 {v[1]} {v[2]} {v[3] or '-'} {float(v[4]):.2f} ← {v[5]}")
    return bad == 0


if __name__ == "__main__":
    install_modules()
    ok = parity()
    ok &= filenames()
    from make_fixture import make
    fx = make(Path(tempfile.mkdtemp()) / "fx")
    ok &= full(fx / "대상.xlsx", fx / "장부", None)
    if os.environ.get("TARGET") and os.environ.get("LEDGERS"):
        ok &= full(Path(os.environ["TARGET"]), Path(os.environ["LEDGERS"]), os.environ.get("BASE"))
    shutdown()
    print("결과:", "모두 같음" if ok else "다른 곳 있음")
    raise SystemExit(0 if ok else 1)
