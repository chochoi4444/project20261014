"""exe(실행 파일) 만들기. Windows에서 실행하면 Windows용 exe가 나온다.

    pip install -r requirements.txt pyinstaller
    python build_exe.py

결과: dist/TotalAssetsFiller.exe      — 더블클릭하면 창이 뜨는 프로그램
      dist/TotalAssetsFiller-cli.exe  — 명령창용 (옵션을 붙여 일괄 처리)
"""
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
COMMON = ["--noconfirm", "--clean", "--onefile", "--paths", str(HERE),
          # 쓰지 않는 큰 라이브러리가 딸려 들어가지 않게
          "--exclude-module", "numpy", "--exclude-module", "pandas", "--exclude-module", "PIL",
          "--exclude-module", "pdfplumber", "--exclude-module", "pytest"]


def build(name, windowed):
    args = [sys.executable, "-m", "PyInstaller", *COMMON, "--name", name, "--windowed" if windowed else "--console",
            str(HERE / "launcher_entry.py")]
    print(">", " ".join(args[2:]))
    subprocess.run(args, check=True, cwd=HERE)


if __name__ == "__main__":
    for d in ("build",):
        shutil.rmtree(HERE / d, ignore_errors=True)
    build("TotalAssetsFiller", windowed=True)
    build("TotalAssetsFiller-cli", windowed=False)
    for spec in HERE.glob("TotalAssetsFiller*.spec"):
        spec.unlink()
    print("완료:", ", ".join(str(p) for p in sorted((HERE / "dist").glob("TotalAssetsFiller*"))))
