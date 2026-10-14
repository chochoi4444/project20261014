# 총자산 자동입력 — 파이썬판

웹 도구(`../src/app.html`)와 **같은 규칙·같은 결과**를 내는 파이썬 명령줄 프로그램입니다. 브라우저 없이 PC나 서버에서 일괄 처리할 때 씁니다.

- 기준가격대장 읽기: `.xlsx` · `.xls` · `.csv` · `.xls`로 저장된 HTML 표
- 펀드명 매칭: 코사인 유사도(글자 1~3gram, TF-IDF) 60% + 글자 순서 일치율 40%, 호수·종류 숫자가 다르면 다른 펀드
- 총자산: `신탁재산총자산` 계정 → `순자산총액 + 부채` → `자산합계` 순, `총자산 = 순자산 + 부채`로 교차검증
- 저장: 원본 xlsx에서 총자산 칸의 XML만 바꿔 **새 파일**로 저장 (서식·데이터 유효성·다른 시트 그대로). 웹판과 결과 파일이 바이트 단위로 같습니다.

## Windows 실행 파일(exe)

파이썬 설치 없이 쓰려면 exe를 받으세요. 저장소 **Releases**의 최신 `총자산 자동입력 exe` → `TotalAssetsFiller-windows.zip`.

- `TotalAssetsFiller.exe` — 더블클릭하면 창이 뜹니다. 대상 엑셀과 장부 폴더를 고르고 ‘실행’.
- `TotalAssetsFiller-cli.exe` — 명령창용 (아래 옵션 동일).
- 사용법은 zip 안의 `README.txt` (= `README_exe.txt`).

exe는 GitHub Actions(`.github/workflows/build-exe.yml`)가 Windows에서 만듭니다. `python/` 아래가 바뀌어 푸시될 때마다 테스트 → exe 빌드 → 가상 장부로 exe 실행 검증 → 릴리스 게시까지 자동으로 합니다.
사내 PC에서 직접 만들려면 Python 3.9 이상을 설치한 뒤 `build_exe.bat`를 실행하면 `dist\`에 exe가 생깁니다.

## 설치

Python 3.9 이상.

```bash
cd nav-total-assets/python
pip install -r requirements.txt
```

## 사용법

```bash
python -m nav_total_assets                      # 창 화면
python -m nav_total_assets 대상.xlsx 장부폴더 --base-date 2026-09-30   # 명령줄
```

| 옵션 | 내용 |
| --- | --- |
| `장부폴더` | 파일이나 폴더를 여러 개 줄 수 있습니다. 폴더는 하위 폴더까지 읽습니다. |
| `--base-date` | 기준일. 생략하면 대상 엑셀의 ‘총자산 (2026.06.30 기준)’ 머리글에서 읽습니다. 장부 기준일이 다르면 ‘검토’가 됩니다. |
| `--include-review` | ‘검토’ 행도 입력합니다. 기본은 ‘자동’과 ‘수동’(매핑으로 지정한) 행만 입력합니다. |
| `--mapping 매핑.json` | 펀드명을 직접 짝지어 줍니다(아래). |
| `--overwrite` | 이미 값이 있는 칸도 덮어씁니다. |
| `-o 결과.xlsx` | 저장 경로. 기본은 `대상이름_총자산입력.xlsx` (원본은 덮어쓰지 않음). |
| `--report 결과.csv` | 매칭 결과 CSV 경로. 기본은 `대상이름_매칭결과.csv`. |
| `--dry-run` | 결과만 보고 엑셀은 저장하지 않습니다. |

실행하면 행별 결과(상태·유사도·총자산·검증·장부 이름)와 요약을 보여 주고, 근거를 담은 CSV를 함께 저장합니다.

**Windows:** 대상 엑셀과 장부 폴더를 `run.bat` 위로 함께 끌어다 놓으면 실행됩니다. 옵션이 필요하면 명령창에서 `run.bat 대상.xlsx 장부폴더 --base-date 2026-09-30`.

### 펀드명 매핑

자동으로 못 찾거나 잘못 찾은 종목은 JSON으로 짝지어 줍니다. 웹 도구의 **매핑 내보내기**로 받은 파일을 그대로 써도 됩니다.

```json
{
  "대상 엑셀의 종목명": "기준가격대장의 펀드명 또는 파일명",
  "매칭하지 않을 종목명": null
}
```

띄어쓰기·괄호 차이는 무시합니다. 매칭 결과 CSV의 ‘매핑키(장부)’ 열 값을 그대로 복사해 쓰면 됩니다.

## 테스트

```bash
pip install pytest
python -m pytest
```

- `test_matching.py` — 웹판과 같은 가상 펀드명 케이스(`../tests/fixtures/matching-cases.json` 공유)
- `test_parity.py` — 웹판 매칭 엔진을 node로 실행해 **모든 점수·근거·배정·상태가 같은지** 대조 (node가 없으면 건너뜀)
- `test_workbook.py` — 총자산 칸만 바꾸는 XML 저장 로직
- `test_synthetic.py` — `tests/make_fixture.py`가 만든 **가상 장부**로 읽기→매칭→저장까지 (실제 데이터 없이 CI에서도 실행)
- 실제 파일로 확인(선택, 저장소 밖 파일):
  `MATCH_CASES=실제케이스.json TARGET=대상.xlsx LEDGERS=장부폴더 BASE_DATE=2026-09-30 python -m pytest -s`
