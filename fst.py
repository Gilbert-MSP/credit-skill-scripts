"""재무구조분석표 엑셀·워드 동시 생성 (분석틀: 하나증권 투자심사1실 박민식 실장, 「재무구조에 대한 이해」)
입력: dart_structure_table 결과 원문(raw/struct.txt). 검토 의견: fst_text.json(5묶음).
사용: python3 kit/fst.py raw/struct.txt --company 회사 --out out [--opinion data/fst_text.json] [--source "자료: ..."]
"""
import argparse, json, os, re
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm
import docx_kit as K

ROWS = [('운전자산', 'W'), ('매입채무', 'P'), ('기타순운전자산', 'O'), ('소계(A) 순영업자산', 'A'),
        ('유형자산', 'U'), ('무형자산', 'I'), ('순투자자산(무형 제외)', 'V'), ('비차입성 장기부채', 'N'),
        ('자기자본(회계상)', 'E'), ('소계(B) 순자기자본여유', 'B'), ('  참고: 무형 차감 자기자본', 'EADJ'),
        ('순필요자금(A+B)', 'AB'), ('금융자산(C)', 'C'), ('구조상 차입금 대응액(A+B+C)', 'ABC'),
        ('단기성차입금', 'DS'), ('장기성차입금', 'DL'), ('총차입금(D)', 'D'), ('검증(A+B+C+D)', 'RES'), ('총자산', 'TA')]
SUBTOTAL = {'A', 'B', 'AB', 'ABC', 'D', 'RES'}
LEAF = ['W', 'P', 'O', 'U', 'I', 'V', 'N', 'E', 'C', 'DS', 'DL', 'TA']
LEAF_LABEL = {lab: code for lab, code in ROWS if code in LEAF}
GREY, FONT = 'F2F4F8', '맑은 고딕'
NUMFMT = '#,##0;(#,##0);"-"'
GROUPS = ["요지", "주요 발견사항", "규명된 핵심 사실", "리스크 요인", "후속 심사 과제"]
NOTE = ("주: 순투자자산은 무형자산 제외·장기부채 상계 전, 자기자본은 비지배지분 포함 회계상 금액이다. 리스부채는 총차입금에 포함했다. "
        "원표의 자산·부채·자본 합계와 분석표 합계가 일치한다. 단수 차이로 합계가 1단위 다를 수 있다.")
FRAME = "분석틀: 하나증권 투자심사1실 박민식 실장, 「재무구조에 대한 이해」"


def period_key(p):
    m = re.match(r'(\d{4})(?:\.(\d{1,2}))?', p)
    return (int(m.group(1)), int(m.group(2) or 12))


def _num(c):
    c = c.strip().replace(",", "")
    if not re.search(r"\d", c):
        return 0.0
    return -float(c.strip("()")) if c.startswith("(") else float(c)


def parse_struct(text: str):
    """→ (기간 오름차순, {코드: [값]}, 단위). 자산 (−), 부채·자본 (+)."""
    unit = (re.search(r"단위:\s*(원|천원|백만원|억원)", text) or [None, "백만원"])[1]
    lines = text.splitlines()
    head = next(l for l in lines if l.startswith("| 구분"))
    periods = [c.strip()[:7] for c in head.strip("|").split("|")[1:]]
    vals = {}
    for l in lines:
        if not l.startswith("|"):
            continue
        cells = [c.strip() for c in l.strip("|").split("|")]
        lab = cells[0].strip("* ")
        if lab in LEAF_LABEL:
            vals[LEAF_LABEL[lab]] = [_num(c) for c in cells[1:1 + len(periods)]]
    missing = [k for k in LEAF if k not in vals]
    if missing:
        raise ValueError(f"구조표에서 행을 찾지 못했다: {missing}")
    order = sorted(range(len(periods)), key=lambda i: period_key(periods[i]))
    return [periods[i] for i in order], {k: [v[i] for i in order] for k, v in vals.items()}, unit


def compute(v, j):
    r = {k: v[k][j] for k in LEAF}
    r['A'] = r['W'] + r['P'] + r['O']
    r['B'] = r['U'] + r['I'] + r['V'] + r['N'] + r['E']
    r['EADJ'] = r['E'] + r['I']
    r['AB'] = r['A'] + r['B']; r['ABC'] = r['AB'] + r['C']
    r['D'] = r['DS'] + r['DL']; r['RES'] = r['ABC'] + r['D']
    return r


def delta_d(calc, i0: int, i1: int) -> dict:
    """ΔD = −(ΔA+ΔB+ΔC) 분해. 합이 맞지 않으면 AssertionError."""
    d = {k: calc[i1][k] - calc[i0][k] for k in ('A', 'B', 'C', 'D', 'W', 'P', 'O', 'U', 'I', 'V', 'N', 'E')}
    assert abs(d['A'] + d['B'] + d['C'] + d['D']) <= 2, d  # 두 기간 각 1단위 반올림
    return d


def fmt(x):
    x = round(x)
    return '-' if x == 0 else (f'({abs(x):,})' if x < 0 else f'{x:,}')


def excel(path, title, asof, unit, periods, vals, note, source):
    wb = Workbook(); ws = wb.active; ws.title = '재무구조분석표'
    ws.sheet_view.showGridLines = False
    base = Font(name=FONT, size=10); bold = Font(name=FONT, size=10, bold=True)
    thin = Side(style='thin', color='BFC5D2')
    ws['A1'] = title; ws['A1'].font = Font(name=FONT, size=14, bold=True, color=K.NAVY)
    ws['A2'] = f"기준일 {asof} · 단위: {unit}"; ws['A2'].font = Font(name=FONT, size=9, color='595959')
    hr = 4
    for c, h in enumerate(['구분'] + periods, 1):
        cell = ws.cell(hr, c, h)
        cell.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=K.NAVY)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    row_of = {code: hr + 1 + i for i, (_, code) in enumerate(ROWS)}
    for label, code in ROWS:
        r = row_of[code]; sub = code in SUBTOTAL
        ws.cell(r, 1, label).font = bold if sub else base
        for j in range(len(periods)):
            L = ws.cell(r, j + 2).column_letter
            ref = lambda k: f'{L}{row_of[k]}'
            f = {'A': f'=SUM({ref("W")}:{ref("O")})', 'B': f'=SUM({ref("U")}:{ref("E")})',
                 'EADJ': f'=SUM({ref("E")},{ref("I")})', 'AB': f'=SUM({ref("A")},{ref("B")})',
                 'ABC': f'=SUM({ref("AB")},{ref("C")})', 'D': f'=SUM({ref("DS")}:{ref("DL")})',
                 'RES': f'=SUM({ref("ABC")},{ref("D")})'}.get(code)
            cell = ws.cell(r, j + 2, f if f else vals[code][j])
            cell.number_format = NUMFMT; cell.font = bold if sub else base
            cell.alignment = Alignment(horizontal='right', vertical='center')
        for c in range(1, len(periods) + 2):
            cell = ws.cell(r, c); cell.border = Border(bottom=thin)
            if sub: cell.fill = PatternFill('solid', fgColor=GREY)
        ws.row_dimensions[r].height = 20
    ws.row_dimensions[hr].height = 22
    last = hr + len(ROWS)
    for k, text in enumerate([note, source]):
        r = last + 2 + k
        ws.cell(r, 1, text).font = Font(name=FONT, size=8, italic=True, color='595959')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(periods) + 1)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[r].height = max(30, 15 * (len(text) // (16 * len(periods)) + 1))
    ws.column_dimensions['A'].width = 32
    for j in range(len(periods)):
        ws.column_dimensions[ws.cell(hr, j + 2).column_letter].width = 17
    ws.freeze_panes = ws.cell(hr + 1, 2)
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def word(path, title, asof, unit, periods, calc, note, source, opinion):
    ncol = len(periods) + 1
    doc = K.new_doc(title, landscape=ncol >= 6)
    sec = doc.sections[0]
    doc.add_paragraph(f"기준일 {asof} · 단위: {unit}")
    size = 8 if ncol <= 7 else 7.5
    t = doc.add_table(rows=1, cols=ncol); t.style = 'Table Grid'; t.autofit = False
    avail = (sec.page_width - sec.left_margin - sec.right_margin) / 36000
    widths = [Mm(58)] + [Mm((avail - 58) / len(periods))] * len(periods)
    for c, h in enumerate(['구분'] + periods):
        cell = t.rows[0].cells[c]; K._shade(cell, K.NAVY); p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; K._font(p.add_run(h), size, True, 'FFFFFF')
    for label, code in ROWS:
        cells = t.add_row().cells; sub = code in SUBTOTAL
        K._font(cells[0].paragraphs[0].add_run(label), size, sub)
        for j, r in enumerate(calc):
            p = cells[j + 1].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            K._font(p.add_run(fmt(r[code])), size, sub)
        if sub:
            for c in cells: K._shade(c, GREY)
    for i, w in enumerate(widths):
        t.columns[i].width = w
        for c in t.columns[i].cells: c.width = w
    K._note_lines(doc, note, source)
    if opinion:
        doc.add_heading("심사역 검토 의견", 2)
        for g in GROUPS:
            doc.add_heading(g + ("(비공개 자료)" if g == "후속 심사 과제" else ""), 3)
            body = opinion.get(g) or []
            if g == "후속 심사 과제":
                body = [f"{x['what']} — {x['why']}. 판단 기준: {x['criterion']}." for x in body] or ["공시 외 자료로 확인할 과제는 없다."]
            for line in ([body] if isinstance(body, str) else body):
                K._font(doc.add_paragraph(style='List Bullet').add_run(line), 9.5)
    doc.save(path)


def build(text: str, company: str, out_dir: str, opinion_path: str | None = None, source: str | None = None,
          basis: str = "연결"):
    opinion = json.load(open(opinion_path, encoding="utf-8")) if opinion_path else None
    if opinion:
        extra = set(opinion) - set(GROUPS)
        if extra:
            raise ValueError(f"검토 의견은 5묶음만 쓴다. 넣을 수 없는 묶음: {sorted(extra)}")
    periods, vals, unit = parse_struct(text)
    calc = [compute(vals, j) for j in range(len(periods))]
    for p, r in zip(periods, calc):
        assert abs(r['RES']) <= 2, f'{p} A+B+C+D = {r["RES"]:,.0f}'
    if "별도" in text.split("\n", 1)[0]:
        basis = "별도"
    src = (source or f"자료: DART 정기보고서({periods[0]}~{periods[-1]}), {basis}")
    src = src if FRAME in src else f"{src}. {FRAME}"
    title = f"{company} {basis} 재무구조분석표"
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.join(out_dir, f"{company}_재무구조분석표_{periods[0]}-{periods[-1]}")
    excel(stem + ".xlsx", title, periods[-1], unit, periods, vals, NOTE, src)
    word(stem + ".docx", title, periods[-1], unit, periods, calc, NOTE, src, opinion)
    return stem + ".xlsx", stem + ".docx"


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("struct"); ap.add_argument("--company", required=True); ap.add_argument("--out", default="out")
    ap.add_argument("--opinion"); ap.add_argument("--source")
    a = ap.parse_args()
    print("saved", *build(open(a.struct, encoding="utf-8").read(), a.company, a.out, a.opinion, a.source))
