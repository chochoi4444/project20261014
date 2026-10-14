"""실행 진입점.

- 인수 없이 실행(exe 더블클릭)하면 창 화면을 연다.
- exe 아이콘 위로 파일·폴더를 끌어다 놓으면 그 경로를 채운 창을 연다 (첫 번째 .xlsx = 대상, 나머지 = 장부).
- 명령줄 인수(옵션)가 있고 콘솔이 있으면 명령줄 모드로 실행한다.
"""
from __future__ import annotations

import sys
from pathlib import Path


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--selftest-gui" in argv:
        from .gui import main as gui_main
        return gui_main(selftest=True)
    has_console = sys.stdout is not None
    dropped_paths_only = bool(argv) and all(not a.startswith("-") and Path(a).exists() for a in argv)
    if argv and has_console and not (dropped_paths_only and len(argv) == 1):
        from .cli import main as cli_main
        return cli_main(argv)
    try:
        from .gui import main as gui_main
    except ImportError:  # tkinter가 없는 환경
        from .cli import main as cli_main
        return cli_main(argv or ["--help"])
    return gui_main(initial_paths=argv if dropped_paths_only else None)
