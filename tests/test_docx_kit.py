"""docx_kit·charts 테스트. 실행: python3 tests/test_docx_kit.py"""
import json, os, re, sys, tempfile, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import docx_kit as K  # noqa: E402
import charts  # noqa: E402
from PIL import Image  # noqa: E402

SRC = "자료: 사업보고서(2025.12), 연결"
M = {"company": "테스트", "unit": "억원", "bases": {"연결": {
    "periods": ["2024.12", "2025.12", "2025.09(누적)", "2026.09(누적)"],
    "src": ["사업보고서 (2025.12)", "사업보고서 (2025.12)",
            "분기보고서 (2025.09)(재무상태표)·분기보고서 (2026.09)(전년동기 손익·현금흐름)", "분기보고서 (2026.09)"],
    "rows": {"매출액": [100, 120, 80, 90], "EBITDA": [20, -5, 10, 12], "EBIT/매출액": [0.1, 0.05, 0.08, 0.09],
             "순차입금(총차입금−금융자산)": [50, -10, 30, 40], "순차입금/EBITDA(배)": [2.5, None, 2.2, 2.4]}}}}


def _xml(doc):
    p = os.path.join(tempfile.mkdtemp(), "t.docx"); doc.save(p)
    return zipfile.ZipFile(p).read("word/document.xml").decode("utf-8")


def _one_table(v=-5):
    d = K.new_doc("제목")
    K.add_table(d, 1, "핵심 지표", "억원", ["구분", "2025"], [["매출액", 100], ["순차입금", v], ["합계", 95]], "주: 테스트", SRC)
    return d


def test_table_requires_source():
    for f in (lambda d: K.add_table(d, 1, "t", "억원", ["a"], [[1]], "주:", ""),
              lambda d: K.add_figure(d, 1, "t", "x.png", "주:", " ")):
        try:
            f(K.new_doc("t"))
        except ValueError:
            continue
        raise AssertionError("출처 없는 표·그림이 들어갔다")


def test_table_followed_by_source():
    x = _xml(_one_table())
    after = x.split("</w:tbl>", 1)[1]
    paras = re.findall(r"<w:p[ >].*?</w:p>", after, re.S)[:2]
    assert any("자료:" in p for p in paras)


def test_margins_font():
    x = _xml(_one_table())
    assert 'w:left="1020"' in x and 'w:top="1020"' in x
    styles = zipfile.ZipFile(_save(_one_table())).read("word/styles.xml").decode("utf-8")
    assert 'w:eastAsia="맑은 고딕"' in styles and '<w:sz w:val="19"/>' in styles


def _save(doc):
    p = os.path.join(tempfile.mkdtemp(), "t.docx"); doc.save(p); return p


def test_negative_burgundy():
    x = _xml(_one_table(-5))
    cell = re.search(r"<w:tc>(?:(?!</w:tc>).)*−5(?:(?!</w:tc>).)*</w:tc>", x, re.S).group(0)
    assert "B8293D" in cell
    assert "n.a." in _xml(_one_table(None))


def test_total_row_shaded():
    x = _xml(_one_table())
    row = [r for r in re.findall(r"<w:tr[ >].*?</w:tr>", x, re.S) if "합계" in r][0]
    assert "EBF0F5" in row


def test_source_line_no_rcept():
    s = K.source_line(M, "연결")
    assert s == "자료: 분기보고서(2025.09), 사업보고서(2025.12), 분기보고서(2026.09), 연결", s
    assert not re.search(r"\d{14}", s)


def test_chart_dpi():
    p = os.path.join(tempfile.mkdtemp(), "c.png")
    charts.sales_ebitda(M, "연결", p)
    assert tuple(round(v) for v in Image.open(p).info["dpi"]) == (300, 300)
    p2 = os.path.join(tempfile.mkdtemp(), "n.png")
    charts.netdebt(M, "연결", p2)
    assert os.path.getsize(p2) > 10000


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
