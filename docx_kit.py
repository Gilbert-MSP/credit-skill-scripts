"""Word 공통: 디자인 상수, 표·그림(출처 없으면 예외), 요약 박스, 자료 줄."""
import re
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

NAVY, GREY, BURGUNDY = "1B365D", "8C99A6", "B8293D"
ALT, TOTAL, BOX = "F8F9FA", "EBF0F5", "F0F4F8"
FONT = "맑은 고딕"
TOTAL_RE = re.compile(r"^\s*(소계|합계|계\b|총계)")


def _font(style_or_run, size=None, bold=None, color=None, italic=None):
    f = style_or_run.font
    f.name = FONT
    el = style_or_run.element
    rpr = el.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)
    if size: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if italic is not None: f.italic = italic
    if color: f.color.rgb = RGBColor.from_string(color)


def _shade(cell, color):
    tcpr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color)
    tcpr.append(sh)


def new_doc(title: str, landscape: bool = False):
    d = Document()
    s = d.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    if landscape:
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(s, side, Mm(18))
    st = d.styles["Normal"]
    _font(st, 9.5)
    st.paragraph_format.line_spacing = 1.25
    st.paragraph_format.space_after = Pt(4)
    for i, sz in ((1, 14), (2, 12), (3, 10.5)):
        _font(d.styles[f"Heading {i}"], sz, True, NAVY)
    d.add_heading(title, 1)
    return d


def fmt(v, kind="num"):
    if v is None:
        return "n.a."
    if isinstance(v, str):
        return v
    if kind == "pct":
        s = f"{abs(v) * 100:,.1f}%"
    elif kind == "x":
        s = f"{abs(v):,.1f}배"
    else:
        s = f"{abs(v):,.0f}"
    return ("−" + s) if v < 0 else s


def _note_lines(doc, note, source):
    if not source or not source.strip():
        raise ValueError("자료: 출처 없이 표·그림을 넣을 수 없다")
    for t in (note, source):
        if t:
            p = doc.add_paragraph()
            r = p.add_run(t if t.startswith(("주:", "자료:")) or t is note else "자료: " + t)
            _font(r, 8, color=GREY)
            p.paragraph_format.space_after = Pt(1)


def add_table(doc, n: int, title: str, unit: str, columns: list, rows: list, note: str, source: str, kinds: list | None = None):
    if not source or not source.strip():
        raise ValueError("자료: 출처 없이 표를 넣을 수 없다")
    cap = doc.add_paragraph()
    _font(cap.add_run(f"[표 {n}] {title}"), 9.5, True, NAVY)
    _font(cap.add_run(f"    (단위: {unit})"), 8, color=GREY)
    cap.paragraph_format.keep_with_next = True
    t = doc.add_table(rows=1, cols=len(columns))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    sec = doc.sections[-1]
    full = (sec.page_width - sec.left_margin - sec.right_margin) / 36000  # mm
    first = min(44, full * 0.4) if len(columns) > 3 else full / len(columns)
    widths = [Mm(first)] + [Mm((full - first) / max(1, len(columns) - 1))] * (len(columns) - 1)
    for j, c in enumerate(columns):
        cell = t.rows[0].cells[j]
        cell.text = ""
        _font(cell.paragraphs[0].add_run(str(c)), 8 if len(columns) > 6 else 9, True, "FFFFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(cell, NAVY)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, w in enumerate(widths):
            cells[j].width = w
        total = isinstance(row[0], str) and bool(TOTAL_RE.match(row[0]))
        kind = (kinds[i] if kinds else "num")
        for j, v in enumerate(row):
            p = cells[j].paragraphs[0]
            txt = str(v) if j == 0 else fmt(v, kind)
            neg = j > 0 and ((isinstance(v, (int, float)) and v < 0) or txt == "EBITDA 적자")
            _font(p.add_run(txt), 9, total, BURGUNDY if neg else None)
            if j > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if total:
                _shade(cells[j], TOTAL)
            elif i % 2:
                _shade(cells[j], ALT)
    for j, w in enumerate(widths):
        t.rows[0].cells[j].width = w
        t.columns[j].width = w  # LibreOffice는 gridCol 너비를 따른다
    _note_lines(doc, note, source)
    return t


def add_figure(doc, n: int, title: str, path: str, note: str, source: str):
    if not source or not source.strip():
        raise ValueError("자료: 출처 없이 그림을 넣을 수 없다")
    cap = doc.add_paragraph()
    _font(cap.add_run(f"[그림 {n}] {title}"), 9.5, True, NAVY)
    doc.add_picture(path, width=Mm(160))
    _note_lines(doc, note, source)


def summary_box(doc, opinion: str, columns: list, rows: list, source: str, kinds: list | None = None):
    if not source or not source.strip():
        raise ValueError("자료: 출처 없이 요약 박스를 넣을 수 없다")
    box = doc.add_table(rows=1, cols=1)
    c = box.rows[0].cells[0]
    _shade(c, BOX)
    tcpr = c._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side, sz, col in (("left", "36", NAVY), ("top", "0", BOX), ("bottom", "0", BOX), ("right", "0", BOX)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single" if sz != "0" else "nil"); e.set(qn("w:sz"), sz); e.set(qn("w:color"), col)
        b.append(e)
    tcpr.append(b)
    p = c.paragraphs[0]
    _font(p.add_run("종합 심사 의견 및 핵심 신용지표 요약"), 10.5, True, NAVY)
    _font(c.add_paragraph().add_run(opinion), 9.5)
    mini = c.add_table(rows=1, cols=len(columns))
    for j, h in enumerate(columns):
        _font(mini.rows[0].cells[j].paragraphs[0].add_run(str(h)), 8.5, True, NAVY)
    for i, row in enumerate(rows):
        cells = mini.add_row().cells
        for j, v in enumerate(row):
            txt = str(v) if j == 0 else fmt(v, kinds[i] if kinds else "num")
            q = cells[j].paragraphs[0]
            _font(q.add_run(txt), 8.5, color=BURGUNDY if (j and isinstance(v, (int, float)) and v < 0) else None)
            if j: q.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _font(c.add_paragraph().add_run(source), 8, color=GREY)


def source_line(metrics: dict, basis: str) -> str:
    seen = []
    for s in metrics["bases"][basis]["src"]:
        for kind, ym in re.findall(r"([가-힣]+보고서)\s*\((\d{4}\.\d{2})\)", s or ""):
            t = f"{kind}({ym})"
            if t not in seen:
                seen.append(t)
    seen.sort(key=lambda t: t[-8:-1])
    return "자료: " + ", ".join(seen) + f", {basis}"
