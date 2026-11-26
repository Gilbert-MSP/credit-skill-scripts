"""light_report.py 테스트. 실행: python3 tests/test_light_report.py (LibreOffice 필요)"""
import copy, json, os, subprocess, sys, tempfile
from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import engine  # noqa: E402
import light_report as L  # noqa: E402

FX = os.path.join(ROOT, "tests", "fixtures")
TEXT = os.path.join(FX, "text_sample.json")
D = tempfile.mkdtemp()
engine.run(os.path.join(FX, "std_sample.txt"), D)
MP = os.path.join(D, "metrics.json")
OUT = os.path.join(D, "light.docx")
ROWS = ["매출액", "EBITDA", "EBITDA 마진", "총차입금", "순차입금(총차입금−금융자산)", "순차입금/EBITDA",
        "EBITDA/금융비용", "부채비율", "운전자본 투자 후 영업현금", "잉여현금흐름(FCF)"]


def _variant(mutate_m=None, mutate_t=None):
    m = json.load(open(MP, encoding="utf-8")); t = json.load(open(TEXT, encoding="utf-8"))
    if mutate_m: mutate_m(m)
    if mutate_t: mutate_t(t)
    d = tempfile.mkdtemp()
    mp, tp, op = (os.path.join(d, x) for x in ("m.json", "t.json", "o.docx"))
    json.dump(m, open(mp, "w", encoding="utf-8"), ensure_ascii=False)
    json.dump(t, open(tp, "w", encoding="utf-8"), ensure_ascii=False)
    return L.build(mp, tp, op), op


def test_build_passes_checks():
    assert L.build(MP, TEXT, OUT) == []


def test_page_count():
    import pypdf
    L.build(MP, TEXT, OUT)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", D, OUT], capture_output=True, timeout=180)
    n = len(pypdf.PdfReader(OUT.replace(".docx", ".pdf")).pages)
    assert 2 <= n <= 3, n


def test_table1_rows():
    L.build(MP, TEXT, OUT)
    t = [x for x in Document(OUT).tables if x.rows[0].cells[0].text == "구분" and len(x.rows) > 8][0]
    assert [r.cells[0].text for r in t.rows[1:]] == ROWS


def test_negative_ebitda_label():
    def neg(m):
        m["bases"]["연결"]["rows"]["EBITDA"][-1] = -10
    issues, op = _variant(neg)
    t = [x for x in Document(op).tables if x.rows[0].cells[0].text == "구분" and len(x.rows) > 8][0]
    row = [r for r in t.rows if r.cells[0].text == "순차입금/EBITDA"][0]
    assert row.cells[-1].text == "EBITDA 적자"


def test_light_separate_only():
    def drop(m):
        del m["bases"]["연결"]
    issues, op = _variant(drop)
    assert "별도" in Document(op).paragraphs[0].text and issues == []


def test_issues_limit():
    def four(t):
        t["issues"] = t["issues"] * 2
    try:
        _variant(mutate_t=four)
    except ValueError:
        return
    raise AssertionError("쟁점 4건이 통과했다")


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
