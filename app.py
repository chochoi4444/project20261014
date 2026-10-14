"""기준가격대장 총자산 취합기 - Streamlit 웹앱

실행: streamlit run app.py
"""
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from aggregator import extract_report, fill_form

SAMPLE = Path(__file__).resolve().parent / "dummy_data"
STATUS_ICON = {"완료": "✅ 완료", "확인필요": "⚠️ 확인필요", "실패": "❌ 실패", "제외": "➖ 제외", "미수령": "⏳ 미수령"}

st.set_page_config(page_title="총자산 취합기", page_icon="📊", layout="wide")
st.title("📊 기준가격대장 총자산 취합기")
st.caption("운용사별 기준가격대장에서 총자산을 찾아 당사 수익증권 평가 양식에 자동으로 채웁니다. "
           "⚠️ 교육용 PoC: 더미데이터만 사용하세요.")

with st.sidebar:
    st.header("1. 파일 올리기")
    report_files = st.file_uploader("운용사 기준가격대장 (여러 개)", type=["xlsx", "xls"], accept_multiple_files=True)
    form_file = st.file_uploader("당사 양식 (총자산 칸이 빈 엑셀)", type=["xlsx"])
    mapping_file = st.file_uploader("종목-운용사 매핑표 (선택)", type=["xlsx"])
    overwrite = st.checkbox("이미 값이 있는 칸도 덮어쓰기", value=False)
    st.divider()
    use_sample = st.button("🧪 샘플 더미데이터로 실행", disabled=not SAMPLE.exists(),
                           help="scripts/make_dummy_data.py 로 만든 파일을 사용합니다.")

if use_sample:
    st.session_state.sample = True
if report_files or form_file:
    st.session_state.sample = False

if st.session_state.get("sample"):
    reports_in = [(p.name, p.read_bytes()) for p in sorted((SAMPLE / "input").glob("*.xls*"))]
    form_in = ("당사양식_수익증권평가_빈양식.xlsx", (SAMPLE / "target" / "당사양식_수익증권평가_빈양식.xlsx").read_bytes())
    mapping_in = (SAMPLE / "master" / "종목_운용사_매핑표.xlsx").read_bytes()
    st.info("샘플 더미데이터로 실행 중입니다.")
else:
    reports_in = [(f.name, f.getvalue()) for f in report_files or []]
    form_in = (form_file.name, form_file.getvalue()) if form_file else None
    mapping_in = mapping_file.getvalue() if mapping_file else None

if not reports_in or not form_in:
    st.markdown("""
#### 사용 방법
1. 왼쪽에서 **기준가격대장 파일들**과 **당사 양식**을 올립니다. 매핑표는 있으면 정확도가 올라갑니다.
2. 앱이 양식(A 좌우형 / B 한 줄 헤더 / C 요약형)을 판별하고 **총자산**을 추출합니다.
3. 종목을 매칭해서 양식에 채우고, 결과를 엑셀로 내려받습니다.

처음이라면 **🧪 샘플 더미데이터로 실행**을 눌러 보세요.
""")
    st.stop()

# ---- 2. 추출
reports = [extract_report(data, name) for name, data in reports_in]
st.subheader("2. 보고서 추출 결과")
st.dataframe(pd.DataFrame([{
    "파일": r.filename, "양식": r.format, "펀드코드": r.fund_code, "펀드명": r.fund_name,
    "기준일": r.base_date, "총자산(원)": r.total_asset, "경고": "; ".join(r.warnings),
} for r in reports]), width="stretch", hide_index=True,
    column_config={"총자산(원)": st.column_config.NumberColumn(format="localized")})

# ---- 3. 기입
try:
    filled, results, unmatched = fill_form(form_in[1], reports, mapping_in, overwrite=overwrite)
except ValueError as e:
    st.error(str(e))
    st.stop()

st.subheader("3. 당사 양식 기입 결과")
counts = pd.Series([r.status for r in results]).value_counts()
cols = st.columns(5)
for col, key in zip(cols, STATUS_ICON):
    col.metric(STATUS_ICON[key], int(counts.get(key, 0)))

status_filter = st.multiselect("상태 필터", list(STATUS_ICON), default=list(STATUS_ICON))
table = pd.DataFrame([{
    "행": r.row, "종목명": r.name, "상태": STATUS_ICON[r.status],
    "총자산(원)": f"{r.total_asset:,}" if isinstance(r.total_asset, int) else (r.total_asset or ""),
    "출처": r.source, "비고": r.remark,
} for r in results if r.status in status_filter])
st.dataframe(table, width="stretch", hide_index=True)

if unmatched:
    st.warning("양식의 어떤 종목과도 연결되지 않은 보고서: " + ", ".join(r.filename for r in unmatched))

st.download_button(
    "⬇️ 총자산이 채워진 양식 내려받기", data=filled, type="primary",
    file_name=f"{Path(form_in[0]).stem}_총자산기입_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
