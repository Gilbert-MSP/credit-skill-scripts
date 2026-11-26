# credit-skill-scripts

신용분석 스킬(credit-review-master, financial-structure-table-report, standard-fs-excel)이 함께 쓰는 공통 스크립트.

## 받기 (스킬이 매 대화 첫 단계에서 실행)

```bash
curl -sL https://github.com/Gilbert-MSP/credit-skill-scripts/archive/refs/heads/main.tar.gz | tar xz && mv credit-skill-scripts-main kit
```

## 파일

| 파일 | 역할 |
|---|---|
| `template/표준재무부표_템플릿_v2.9.xlsx` | 표준재무부표 양식. 지표 산식의 유일한 기준 |
| `stdfs_fill.py` | `dart_std_data` 결과를 양식에 채운다 |
| `engine.py` | 양식 채우기 → LibreOffice 재계산 → `metrics.json`·`qa.json` |
| `docx_kit.py` · `charts.py` | Word 서식·출처 줄·필수 차트 |
| `light_report.py` | 약식 점검 Word 2~3쪽 |
| `fst.py` | 재무구조분석표 엑셀·워드 |
| `checks.py` | 출고 전 검사(문체·출처·숙제 문구·수치 대사) |

테스트: `for t in tests/test_*.py; do python3 $t; done`

## 주의

공개 저장소다. API 키, 회사 분석 결과, 사내 자료를 올리지 않는다. 테스트 표본은 공시된 숫자만 쓴다.

분석틀 출처: 하나증권 투자심사1실 박민식 실장
