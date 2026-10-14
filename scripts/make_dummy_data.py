"""
AI 바이브코딩 교육용 더미데이터 생성기 (과제: 비정형 보고서 취합 에이전트)

- 실제 회사/펀드/금액은 하나도 쓰지 않습니다. 모든 이름과 숫자는 가상입니다.
- 실행: python scripts/make_dummy_data.py   (필요 패키지: openpyxl)
- 결과: 저장소 루트의 dummy_data/ 폴더에 아래 파일 생성
    input/  운용사별 기준가격대장 (양식 A/B/C 3종, 일부러 서로 다르게 생성)
    master/ 종목_운용사_매핑표.xlsx
    target/ 당사양식_수익증권평가_빈양식.xlsx   (총자산 칸이 비어 있음)
    answer/ 당사양식_정답.xlsx                  (검증용 기대값)
"""
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook

random.seed(2026)  # 매번 같은 더미데이터가 나오도록 고정
BASE_DATE = date(2026, 6, 30)
OUT = Path(__file__).resolve().parents[1] / "dummy_data"

MANAGERS = {
    "A": "알파자산운용",
    "B": "베타인베스트먼트",
    "C": "감마파트너스",
}

# (종목명, 표준코드, 펀드코드, 운용사양식, 유가증권분류, 구조화기업성격, 운용펀드)
FUNDS = [
    ("알파 글로벌크레딧 사모특별자산 1호", "KRZZ00000011", "A10001", "A", "기타수익증권", "투자펀드 및 투자신탁", "대체-시장성-특별자산"),
    ("알파 미국인프라 사모투자신탁 2호",   "KRZZ00000029", "A10002", "A", "기타수익증권", "투자펀드 및 투자신탁", "대체-투자금융-SOC"),
    ("알파 유럽부동산 재간접 3호",         "KRZZ00000037", "A10003", "A", "기타수익증권", "투자펀드 및 투자신탁", "대체-투자금융-부동산"),
    ("베타 CLO 레버리지 노트 7호",         "KRZZ00000045", "B20007", "B", "기타수익증권", "자산유동화 SPC",       "대체-시장성-특별자산"),
    ("베타 선순위대출 사모펀드 8호",       "KRZZ00000052", "B20008", "B", "기타수익증권", "투자펀드 및 투자신탁", "대체-시장성-특별자산"),
    ("베타 항공기금융 사모 9호",           "KRZZ00000060", "B20009", "B", "기타수익증권", "선박금융 및 인수금융", "대체-투자금융-SOC"),
    ("감마 신재생에너지 사모특별자산 1호", "KRZZ00000078", "C30001", "C", "주식형",       "부동산금융",           "대체-투자금융-SOC"),
    ("감마 물류센터 PFV 사모 2호",         "KRZZ00000086", "C30002", "C", "기타수익증권", "부동산금융",           "대체-투자금융-부동산"),
    ("(주)가상테크 상장주식",              "KRZZ00000094", "-",      "-", "코스닥주식",   "해당없음",             "국내주식-순수"),  # 제외 규칙 테스트용(상장주식)
    ("알파 단기채 MMF 1호",                "KRZZ00000110", "-",      "-", "MMF",          "MMF",                  "단기금융"),       # 제외 규칙 테스트용(MMF)
    ("감마 세컨더리 PEF 3호",              "KRZZ00000102", "C30003", "C", "PEF",          "투자펀드 및 투자신탁", "대체-투자금융-PEF"),  # 보고서 미수령 테스트용
]
MISSING_REPORT = {"C30003"}  # 운용사 보고서가 아직 안 온 펀드


def make_ledger_numbers():
    """기준가격대장 핵심 숫자를 서로 맞게(검증식 성립) 생성"""
    deposit = random.randint(5, 50) * 10_000_000 + random.randint(0, 9_999_999)
    invest = random.randint(30, 500) * 100_000_000 + random.randint(0, 99_999_999)
    accrued = random.randint(1_000_000, 900_000_000)
    total_asset = deposit + invest + accrued
    liab = random.randint(5_000_000, 150_000_000)
    nav = total_asset - liab
    units = random.randint(30, 500) * 100_000_000
    price = round(nav / units * 1000, 2)  # 1,000좌당 기준가
    return dict(deposit=deposit, invest=invest, accrued=accrued, total_asset=total_asset,
                liab=liab, mgmt_fee=int(liab * 0.8), trustee_fee=liab - int(liab * 0.8),
                principal=units, nav=nav, units=units, price=price,
                prev_total=total_asset - random.randint(-50_000_000, 50_000_000))


def write_format_a(path, name, code, n):
    """양식 A: 차변/대변 좌우 배치 + 상단에 작업일자·펀드명 (요약코드 100000/200000/7000xx)"""
    wb = Workbook(); ws = wb.active; ws.title = "Page"
    ws["B2"] = "기준가격대장"; ws["F2"] = "작업일자 ▶"; ws["H2"] = BASE_DATE.strftime("%Y-%m-%d")
    ws["F3"] = "펀 드 명 ▶"; ws["H3"] = code; ws["I3"] = name
    ws["B4"] = f"{MANAGERS['A']} (가상)"
    ws.append([]); ws.append(["차   변", None, None, None, "계 정 과 목", None, "대    변"])
    ws.append(["당일잔액", "당일증감액", "전일잔액", None, "코드", "계정명", "전일잔액", "당일증감액", "당일잔액"])
    rows = [
        (n["total_asset"], n["total_asset"] - n["prev_total"], n["prev_total"], "100000", "자산합계", None, None, None),
        (n["deposit"], None, n["deposit"], "101010", "예금", None, None, None),
        (n["invest"], None, n["invest"], "161020", "대여금(금전채권)", None, None, None),
        (n["accrued"], None, n["accrued"], "149510", "대여금미수이자", None, None, None),
        (None, None, None, "200000", "부채합계", n["liab"], None, n["liab"]),
        (None, None, None, "283011", "미지급운용보수", n["mgmt_fee"], None, n["mgmt_fee"]),
        (None, None, None, "281011", "미지급수탁보수", n["trustee_fee"], None, n["trustee_fee"]),
        (None, None, None, "300000", "원본합계", n["principal"], None, n["principal"]),
        (None, None, None, "700001", "순자산총액", None, None, n["nav"]),
        (None, None, None, "700003", "좌수", None, None, n["units"]),
        (None, None, None, "700004", "기준가", None, None, n["price"]),
    ]
    for d1, d2, d3, c, nm, c1, c2, c3 in rows:
        ws.append([d1, d2, d3, None, c, nm, c1, c2, c3])
    wb.save(path)


def write_format_b(path, name, code, n):
    """양식 B: 한 줄 헤더 표 + 계층형 계정명([[ 자산 ]]) + 10자리 계정코드, 펀드명은 파일명에만 있음"""
    wb = Workbook(); ws = wb.active; ws.title = "10701_기준가격대장 조회"
    ws.append(["계정구분", "당일차변잔액", "차변금액", "전일차변잔액", "계정코드", "계정과목명", "전일대변잔액", "대변금액", "당일대변잔액"])
    ws.append([None, None, None, None, "0000000000", "계정구분 : A"])
    ws.append([None, n["total_asset"], n["total_asset"] - n["prev_total"], n["prev_total"], "1000000000", "[[ 자산 ]]", 0, 0, 0])
    ws.append([None, n["deposit"], 0, n["deposit"], "1010100000", "  예금(은대)", 0, 0, 0])
    ws.append([None, n["invest"], 0, n["invest"], "1061300000", "  해외채권", 0, 0, 0])
    ws.append([None, n["accrued"], 0, n["accrued"], "1070100000", "  단기상품미수이자", 0, 0, 0])
    ws.append([None, 0, 0, 0, "2000000000", "[[ 부채 ]]", n["liab"], 0, n["liab"]])
    ws.append([None, 0, 0, 0, "4010100000", "원본액", n["principal"], 0, n["principal"]])
    ws.append([None, None, None, None, "B016100000", "순자산", None, None, n["nav"]])
    ws.append([None, None, None, None, "9010500000", "[ 수익증권잔존좌수 ]", None, None, n["units"]])
    ws.append([None, None, None, None, "B010100000", "기준가격", None, None, n["price"]])
    ws.append([None, None, None, None, "0000000000", "합계 : 11 건"])
    wb.save(path)


def write_format_c(path, name, code, n):
    """양식 C: 운용보고서 요약형(항목/값 세로 나열, 단위 '백만원', 날짜 표기 다름)"""
    wb = Workbook(); ws = wb.active; ws.title = "요약"
    ws.append([f"{name} 분기 운용현황 보고"])
    ws.append([f"기준일 : {BASE_DATE.year}년 {BASE_DATE.month}월 {BASE_DATE.day}일", None, "(단위: 백만원)"])
    ws.append([f"운용사 : {MANAGERS['C']} (가상)", None, f"펀드코드 : {code}"])
    ws.append([])
    ws.append(["구분", "금액"])
    to_m = lambda v: round(v / 1_000_000, 1)
    ws.append(["총자산(AUM)", to_m(n["total_asset"])])
    ws.append(["총부채", to_m(n["liab"])])
    ws.append(["순자산(NAV)", to_m(n["nav"])])
    commitment = random.randint(0, 30) * 100_000_000
    ws.append(["잔여약정금액", to_m(commitment)])
    ws.append(["기준가(원)", n["price"]])
    wb.save(path)
    return commitment


TARGET_HEADER = ["기준일", "포트폴리오", "G/L 계정명", "종목명", "표준코드", "거래상대방", "운용펀드",
                 "유가증권분류", "발행일", "만기일", "기초장부금액(원화)", "구조화기업성격",
                 f"총자산({BASE_DATE:%Y.%m.%d} 기준)(원)", "잔여약정금액(원)", "운용부서", "비고"]


def main():
    for sub in ["input", "master", "target", "answer"]:
        (OUT / sub).mkdir(parents=True, exist_ok=True)

    master = Workbook(); ms = master.active; ms.title = "매핑표"
    ms.append(["종목명", "표준코드", "펀드코드", "운용사", "보고서양식", "파일명"])
    target = Workbook(); ts = target.active; ts.title = "수익증권"; ts.append(TARGET_HEADER)
    answer = Workbook(); ans = answer.active; ans.title = "정답"; ans.append(TARGET_HEADER)

    for name, std, code, fmt, cls, nature, fund_type in FUNDS:
        filename, commitment = "", None
        if fmt == "-":                               # 상장주식·MMF → 총자산 '제외'
            total, remark = "제외", f"제외 대상({cls})"
        elif code in MISSING_REPORT:                 # 보고서 미수령
            total, remark = None, "보고서 미수령"
            filename = f"{BASE_DATE:%Y%m%d}_{code}_기준가격대장.xlsx"
        else:
            n = make_ledger_numbers()
            mgr = MANAGERS[fmt]
            if fmt == "A":
                filename = f"{BASE_DATE:%Y%m%d}_{mgr}_{name}_기준가격대장.xlsx"
                write_format_a(OUT / "input" / filename, name, code, n)
            elif fmt == "B":
                filename = f"{mgr}_{name.replace(' ', '')}_{BASE_DATE:%y%m%d}.xlsx"
                write_format_b(OUT / "input" / filename, name, code, n)
            else:
                filename = f"{name}_운용보고서_{BASE_DATE.year}_2Q.xlsx"
                commitment = write_format_c(OUT / "input" / filename, name, code, n)
            # 양식 C는 백만원 단위라 원 단위로 복원 시 반올림 오차가 생김 → 정답도 같은 방식으로
            total = round(n["total_asset"] / 1_000_000, 1) * 1_000_000 if fmt == "C" else n["total_asset"]
            total, remark = int(round(total)), ""

        ms.append([name, std, code, MANAGERS.get(fmt, "-"), fmt, filename])
        row = [BASE_DATE, "일반계정", f"FVPL-{cls}(NT)", name, std, MANAGERS.get(fmt, "가상증권(주)"),
               fund_type, cls, date(2021, random.randint(1, 12), 1), date(2031, random.randint(1, 12), 1),
               random.randint(5, 300) * 100_000_000, nature]
        ts.append(row + [None, None, "투자금융팀", None])
        ans.append(row + [total, None, "투자금융팀", remark])  # 잔여약정금액은 PoC 대상 아님(양식 C에만 방해 항목으로 존재)

    master.save(OUT / "master" / "종목_운용사_매핑표.xlsx")
    target.save(OUT / "target" / "당사양식_수익증권평가_빈양식.xlsx")
    answer.save(OUT / "answer" / "당사양식_정답.xlsx")
    print(f"생성 완료 → {OUT}")
    for p in sorted(OUT.rglob("*.xlsx")):
        print(" -", p.relative_to(OUT))


if __name__ == "__main__":
    main()
