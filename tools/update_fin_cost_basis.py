"""(한 번만 쓰는 도구) 평가사 기준 금융비용(T46)을 평가사 부표 정의에 맞춘다.
  한국기업평가 부표: 이자비용 + 매출채권처분손실 + 자본화된 이자비용
  NICE신용평가 부표: 이자비용 + 매출채권처분손실 (자본화 이자 없음)
→ T46 = 이자비용 + 신종자본증권 배당(설정 B6) + 매출채권처분손실 + 설정 B12 × 자본화 이자. 기타 금융원가·보증비용은 넣지 않는다.
사용: python3 tools/update_fin_cost_basis.py template/표준재무부표_템플릿_v2.9.xlsx"""
import copy
import sys

import openpyxl

C = "CDEFGHI"
OLD = "참고: 금융비용(평가사 기준: 이자비용+매출채권처분손실+보증비용+자본화 이자)"
NEW = "참고: 금융비용(평가사 기준: 이자비용+매출채권처분손실+자본화 이자(설정))"

wb = openpyxl.load_workbook(sys.argv[1])
st = wb["설정"]
for j in range(1, 4):
    st.cell(12, j)._style = copy.copy(st.cell(11, j)._style)
st["A12"], st["B12"] = "자본화 이자를 평가사 기준 금융비용에 포함", 1
st["C12"] = "1 = 한국기업평가 부표, 0 = NICE신용평가 부표(이자비용+매출채권처분손실만)"
for basis in ("연결", "별도"):
    ws = wb[f"표준부표({basis})"]
    at = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
    ws.cell(at["R20"] - 1, 2, "Ⅸ. 참고: 평가사 기준 금융비용 (한기평 부표: 이자비용+매출채권처분손실+자본화 이자, NICE 부표: 자본화 이자 제외 → 설정 B12. 보증비용은 합계에 넣지 않음)")
    for c in C:
        ws[f"{c}{at['T46']}"] = f"={c}{at['R9']}+설정!$B$6*{c}{at['R19']}+N({c}{at['R20']})+설정!$B$12*N({c}{at['R22']})"
    mi = wb[f"주요 재무지표({basis})"]
    for r in range(1, mi.max_row + 1):
        if mi.cell(r, 1).value == OLD:
            mi.cell(r, 1, NEW)
wb.save(sys.argv[1])
print("saved", sys.argv[1])
