"""(한 번만 쓰는 도구) 평가사 기준 금융비용 줄 추가.
표준부표(연결·별도) 맨 아래에 입력 코드 R20·R21·R22와 산식 T46~T48을 붙이고(기존 행 번호는 그대로),
주요 재무지표 시트의 '이자보상배율' 아래에 참고 줄 3개를 끼워 넣는다.
사용: python3 tools/add_fin_cost_rows.py template/표준재무부표_템플릿_v2.9.xlsx"""
import copy
import sys

import openpyxl

C = "CDEFGHI"          # 표준부표 기간 칸
BASE = "(C78+C79+설정!$B$6*C88)"   # 지금 양식의 금융비용 = 이자비용 + 기타 금융원가 + 신종자본증권 배당(설정)


def style(dst, src):
    if src.has_style:
        dst._style = copy.copy(src._style)


def add_codes(ws):
    r0 = ws.max_row + 2
    rows = [("", "Ⅸ. 참고: 평가사 기준 금융비용 (NICE신용평가 정의: 손익계산서 밖 자본화 이자 포함)", None, 92),
            ("R20", "매출채권처분손실", None, 78),
            ("R21", "금융보증·지급보증 비용", None, 78),
            ("R22", "자본화된 차입원가(손익계산서 밖)", None, 78),
            ("T46", "금융비용(평가사 기준)", "={b}+N({c}{r1})+N({c}{r2})+N({c}{r3})", 78),
            ("T47", "EBITDA/금융비용(평가사 기준, 배)", '=IFERROR(IF({c}{t}<1,"n.m.",{c}90/{c}{t}),"")', 173),
            ("T48", "EBIT/금융비용(평가사 기준, 배)", '=IFERROR(IF({c}{t}<1,"n.m.",{c}89/{c}{t}),"")', 173)]
    at = {}
    for i, (code, name, f, srow) in enumerate(rows):
        r = r0 + i
        at[code] = r
        for j in range(1, 10):
            style(ws.cell(r, j), ws.cell(srow, j))
        ws.cell(r, 1, code or None)
        ws.cell(r, 2, name)
    for code, name, f, srow in rows:
        if not f:
            continue
        for c in C:
            ws[f"{c}{at[code]}"] = f.format(b=BASE.replace("C", c).replace("$B$6", "$B$6"), c=c, r1=at["R20"], r2=at["R21"], r3=at["R22"], t=at["T46"])
    return at


def add_indicators(ws, src: str, at: dict):
    lab = "이자보상배율(EBIT/금융비용, 배)"
    r = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, 1).value == lab)
    new = [("참고: 금융비용(평가사 기준: 이자비용+매출채권처분손실+보증비용+자본화 이자)", "T46", r - 3 if False else None),
           ("참고: EBITDA/금융비용(평가사 기준, 배)", "T47", None),
           ("참고: EBIT/금융비용(평가사 기준, 배)", "T48", None)]
    amount_row = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, 1).value == "총차입금")
    for k, (name, code, _) in enumerate(new):
        ws.insert_rows(r + 1 + k)
        rr = r + 1 + k
        s = amount_row if code == "T46" else r
        for j in range(1, 9):
            style(ws.cell(rr, j), ws.cell(s if s < rr else s + 1, j))
        ws.cell(rr, 1, name)
        for j, c in enumerate("BCDEFGH"):
            ws[f"{c}{rr}"] = f"=IF('{src}'!{C[j]}{at[code]}=\"\",\"\",'{src}'!{C[j]}{at[code]})"


if __name__ == "__main__":
    wb = openpyxl.load_workbook(sys.argv[1])
    for basis in ("연결", "별도"):
        src = f"표준부표({basis})"
        at = add_codes(wb[src])
        add_indicators(wb[f"주요 재무지표({basis})"], src, at)
    wb.save(sys.argv[1])
    print("saved", sys.argv[1])
