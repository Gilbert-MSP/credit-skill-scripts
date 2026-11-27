"""(한 번만 쓰는 도구) '주요 재무지표' 시트에 계약자산·계약부채·파생상품 줄을 끼워 넣는다. 기존 줄의 수식은 건드리지 않는다.
사용: python3 tools/add_indicator_rows.py template/표준재무부표_템플릿_v2.9.xlsx"""
import copy
import sys

import openpyxl

COLS = "BCDEFGH"            # 지표 시트 B~H = 표준부표 C~I (7개 기간)


def src_col(c: str) -> str:
    return chr(ord(c) + 1)


def find(ws, label: str) -> int:
    return next(r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip() == label.strip())


def insert(ws, after: int, label: str, expr: str, style_row: int):
    """after 줄 바로 아래에 새 줄. expr의 {S}는 표준부표 시트 이름, {c}는 표준부표 열 글자"""
    ws.insert_rows(after + 1)
    r = after + 1
    s = style_row + 1 if style_row > after else style_row
    for j in range(1, 9):
        src, dst = ws.cell(s, j), ws.cell(r, j)
        if src.has_style:
            dst._style = copy.copy(src._style)
    ws.cell(r, 1, label)
    for c in COLS:
        ws[f"{c}{r}"] = expr.format(S=ws.title.replace("주요 재무지표", "'표준부표") + "'", c=src_col(c))
    return r


def apply(wb):
    for basis in ("연결", "별도"):
        ws = wb[f"주요 재무지표({basis})"]
        S = "{S}"
        # 아래 줄부터 끼워 넣어야 위쪽 줄 번호가 그대로다
        r = find(ws, "매출채권·계약자산 회전기간(일)")
        insert(ws, r, "계약자산/매출액", '=IFERROR(IF(' + S + '!{c}70<=0,"n.m.",' + S + '!{c}17/(' + S + '!{c}70*' + S + '!{c}9)),"")', r)
        r = find(ws, "단기성차입금")
        r = next(x for x in range(r + 1, ws.max_row + 1) if ws.cell(x, 1).value and ws.cell(x, 1).value.strip() == "단기성차입금")
        r2 = insert(ws, r, "참고: 파생상품부채 포함 총차입금(전환권 등 상환 의무가 있을 때만 의미)",
                    '=IFERROR(' + S + '!{c}151+(1-설정!$B$5)*' + S + '!{c}51,"")', r)
        r3 = insert(ws, r2, "   파생상품부채", '=IFERROR(' + S + '!{c}51,"")', find(ws, "   그중 리스부채"))
        insert(ws, r3, "   파생상품자산", '=IFERROR(' + S + '!{c}30,"")', find(ws, "   그중 리스부채"))
        r = find(ws, "   그중 자본성 증권")
        insert(ws, r, "   그중 파생상품부채('설정' B5 = 1일 때)", '=IFERROR(설정!$B$5*' + S + '!{c}51,"")', r)
        r = next(x for x in range(1, ws.max_row + 1) if str(ws.cell(x, 1).value or "").startswith("운전자금(매출채권+계약자산"))
        ws.cell(r, 1, "수주산업 운전자금(매출채권+계약자산+재고−매입채무−계약부채)")
        insert(ws, r, "순수 운전자금(매출채권+재고−매입채무)", '=IFERROR(' + S + '!{c}16+' + S + '!{c}18-' + S + '!{c}38,"")', r)
        r = find(ws, "   계약부채·선수금")
        ws.cell(r, 1, "   계약부채(선수금 포함)")
        r = find(ws, "   매출채권·계약자산")
        r2 = insert(ws, r, "      그중 매출채권", '=IFERROR(' + S + '!{c}16,"")', r)
        insert(ws, r2, "      그중 계약자산", '=IFERROR(' + S + '!{c}17,"")', r)


if __name__ == "__main__":
    wb = openpyxl.load_workbook(sys.argv[1])
    apply(wb)
    wb.save(sys.argv[1])
    print("saved", sys.argv[1])
