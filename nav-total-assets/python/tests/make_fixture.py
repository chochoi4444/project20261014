"""가상의 대상 엑셀과 기준가격대장을 만든다 (실제 데이터 없이 빌드·테스트하기 위해).

    python tests/make_fixture.py 출력폴더

대상.xlsx 1개와 장부 2개(시산표형 두 양식)를 만들고, 기대하는 총자산을 expected.json에 적는다.
"""
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation


def make(out: Path):
    out.mkdir(parents=True, exist_ok=True)
    ledgers = out / "장부"
    ledgers.mkdir(exist_ok=True)

    # 대상 엑셀: 머리글에 기준일, MMF는 '제외', 서식과 데이터 유효성 포함
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "1. 수익증권 주식"
    ws.append(["key", "종목명", "유가증권분류", "총자산\n(2026.09.30 기준) (원)", "표준코드"])
    rows = [
        ["", "가온대체투자전문투자형사모부동산투자신탁101호", "기타수익증권", None, "KRZ000000001"],
        ["", "누리파트너스전문투자형사모부동산투자신탁유럽제1호", "기타수익증권", None, "KRZ000000002"],
        ["", "가온대체투자전문투자형사모부동산투자신탁102호", "기타수익증권", None, "KRZ000000003"],
        ["", "예시법인용MMF 1(국공채)C", "MMF", "제외", "KRZ000000004"],
        ["", "보람에너지풍력사모특별자산투자신탁", "주식형", None, "KRZ000000005"],
        ["", "보람에너지태양광사모특별자산투자신탁", "주식형", None, "KRZ000000006"],
        ["", "보람에너지지열사모특별자산투자신탁", "주식형", None, "KRZ000000007"],
    ]
    for r in rows:
        ws.append(r)
    for r in range(2, 9):
        ws.cell(r, 4).number_format = "#,##0"
    dv = DataValidation(type="list", formula1='"A,B"')
    ws.add_data_validation(dv)
    dv.add("F2:F5")
    wb.save(out / "대상.xlsx")

    # 양식 B형: 10자리 계정코드, 당일차변잔액/당일대변잔액, 본문에 펀드명 없음(파일명 사용)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["계정구분", "당일차변잔액", "차변금액", "전일차변잔액", "계정코드", "계정과목명", "전일대변잔액", "대변금액", "당일대변잔액"])
    ws.append([None, 5_100_000_000, 0, 5_100_000_000, "1000000000", "[[ 자산 ]]", 0, 0, 0])
    ws.append([None, 0, 0, 0, "2000000000", "[[ 부채 ]]", 10_000_000, 0, 12_345_678])
    ws.append([None, 4_987_654_322, 0, 4_987_654_322, "9010400000", "[ 신탁재산순자산총액 ]", 0, 0, 0])
    ws.append([None, 1012.34, 0, 1012.34, "9010600000", "수익증권기준가격", 0, 0, 0])
    ws.append([None, 5_000_000_000, 0, 5_000_000_000, "G010100000", "신탁재산총자산", 0, 0, 0])
    wb.save(ledgers / "가온대체투자일반사모부동산투자신탁101호_기준가격대장_260930(결산후).xlsx")
    # 같은 이름의 결산전 파일: 결산후 파일이 있으므로 쓰지 않아야 한다
    ws["B2"] = 4_900_000_000
    wb.save(ledgers / "가온대체투자일반사모부동산투자신탁101호_기준가격대장_260930(결산전).xlsx")

    # 양식 C형: 6자리 계정코드, 머리글에 펀드명·운용사, 총자산 계정 없음 → 순자산총액 + 부채합계
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([None, "기준가격대장", None, None, None, "작업일자 ▶", None, "2026.09.30"])
    ws.append([None, None, None, None, None, "펀 드 명 ▶", None, 123456, "누리파트너스해외부동산투자신탁1호(전문)"])
    ws.append([None, "누리파트너스자산운용"])
    ws.append([])
    ws.append(["당일잔액", None, "당일증감액", "전일잔액", None, "코드", "계정명", None, None, "전일잔액", None, "당일증감액", None, "당일잔액"])
    ws.append([21_000_000_000, None, 0, 21_000_000_000, None, "100000", "자산합계"])
    ws.append([None, None, None, None, None, "200000", "부채합계", None, None, 400_000_000, None, 100_000, None, 400_100_000])
    ws.append([None, None, None, None, None, "700001", "순자산총액", None, None, 20_600_000_000, None, -100_000, None, 20_599_900_000])
    wb.save(ledgers / "누리해외부동산1호_신탁재산명세부_260930.xlsx")

    # 통합 파일: 요약 시트 + 펀드마다 '총계정원장_' 시트. 머리글은 '금일잔액', 펀드명은 '펀드 :' 칸 뒤 코드·이름
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "마감자료"
    ws.append(["구분", "보람에너지풍력", "보람에너지태양광"])
    ws.append(["NAV", 2_990_000_000, 1_495_000_000])
    for sheet, code, name, asset, liab in [("총계정원장_풍력", 10107, "보람에너지풍력발전사모특별자산투자신탁[인프라사업]", 3_000_000_000, 10_000_000),
                                           ("총계정원장_태양광", 10111, "보람에너지태양광사모특별자산투자신탁[인프라사업]", 1_500_000_000, 5_000_000)]:
        ws = wb.create_sheet(sheet)
        ws.append([None, None, None, "총 계 정 원 장"])
        ws.append([None, None, None, "기준일 : 2026.09.30 (일반)"])
        ws.append(["펀드 :", code, name, None, None, None, None, "보람자산운용"])
        ws.append([])
        ws.append(["차변", None, None, "계정명", "계정코드", "대변"])
        ws.append(["금일잔액", "금일증감", "전일잔액", None, None, "전일잔액", "금일증감", "금일잔액"])
        ws.append([asset, 0, asset, "[자산]", "1000000"])
        ws.append([None, None, None, "[부채]", "2000000", liab, 0, liab])
        ws.append([None, None, None, "순자산총액", "9011100", asset - liab, 0, asset - liab])
    # 보유종목 명세 탭(원장 아님)과, 펀드명 칸 없이 제목 아래 엑셀 날짜 숫자만 있는 원장 탭
    ws = wb.create_sheet("통합명세부_풍력")
    ws.append([None, "통합명세부"])
    ws.append(["자산구분", "종목명", "수량", "취득액", "평가손익"])
    ws.append(["주식", "가상발전(주)", 100, 2_900_000_000, 0])
    ws = wb.create_sheet("총계정원장_보람에너지지열")
    ws.append([None, None, None, "총계정원장"])
    ws.append([None, None, None, 46295])
    ws.append([])
    ws.append(["금일잔액", "금일증감", "전일잔액", "계정과목", "계정코드", "전일잔액", "금일증감", "금일잔액"])
    ws.append([800_000_000, 0, 800_000_000, "[자산]", "100000"])
    ws.append([None, None, None, "[부채]", "200000", 2_000_000, 0, 2_000_000])
    ws.append([None, None, None, "신탁재산순자산총액", "A10300", 798_000_000, 0, 798_000_000])
    wb.save(ledgers / "02. 보람에너지 마감자료_260930.xlsx")

    expected = {"가온대체투자전문투자형사모부동산투자신탁101호": 5_100_000_000,
                "누리파트너스전문투자형사모부동산투자신탁유럽제1호": 21_000_000_000}
    (out / "expected.json").write_text(json.dumps(expected, ensure_ascii=False), encoding="utf-8")
    return out


if __name__ == "__main__":
    print(make(Path(sys.argv[1] if len(sys.argv) > 1 else "fixture_out")))
