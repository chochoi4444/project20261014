# PyInstaller 진입 스크립트 (패키지 밖에 있어야 상대 import가 동작한다)
from nav_total_assets.launcher import main

if __name__ == "__main__":
    raise SystemExit(main())
