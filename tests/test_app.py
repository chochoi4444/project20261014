"""Streamlit 화면이 샘플 데이터로 오류 없이 렌더링되는지 확인."""
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parents[1] / "app.py")


def test_app_runs_with_sample_data():
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    at.sidebar.button[0].click().run()
    assert not at.exception
    assert at.metric[0].value == "8"  # ✅ 완료 8건
