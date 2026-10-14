import json
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
FIXTURE = HERE.parent.parent / "tests" / "fixtures" / "matching-cases.json"


@pytest.fixture(scope="session")
def cases():
    return json.loads(FIXTURE.read_text(encoding="utf-8"))
