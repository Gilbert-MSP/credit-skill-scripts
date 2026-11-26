"""fst.py 테스트. 실행: python3 tests/test_fst.py"""
import json, os, sys, tempfile
import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import fst  # noqa: E402
import checks  # noqa: E402

FX = os.path.join(ROOT, "tests", "fixtures")
TEXT = open(os.path.join(FX, "struct_sample.txt"), encoding="utf-8").read()
OPINION = {
    "요지": "총차입금은 2,830억원으로 구조상 차입금 대응액과 일치하며 자금 조달 구조는 안정적이라고 판단된다.",
    "주요 발견사항": ["순영업자산 소요가 448억원 늘었으나 자기자본 증가로 충당되었다."],
    "규명된 핵심 사실": ["9개월간 총차입금 증가 25억원은 운전자산 증가 402억원과 금융자산 감소 41억원의 합성 결과이다."],
    "리스크 요인": ["장기성차입금 비중이 76.8%로 높아져 차환 시점이 2년 이후로 몰려 있다."],
    "후속 심사 과제": [{"what": "회사채 인수 약정서의 재무약정", "why": "사채 1,795억원", "criterion": "부채비율 한도가 있으면 여유 폭을 계산한다"}],
}
D = tempfile.mkdtemp()
OP = os.path.join(D, "op.json")
json.dump(OPINION, open(OP, "w", encoding="utf-8"), ensure_ascii=False)
XLSX, DOCX = fst.build(TEXT, "한솔케미칼", D, OP)


def test_identity():
    periods, vals, unit = fst.parse_struct(TEXT)
    for j, p in enumerate(periods):
        r = fst.compute(vals, j)
        assert abs(r["RES"]) <= 1, (p, r["RES"])  # 억원 반올림 단수 차이 1단위 허용
    assert unit == "억원" and vals["D" if "D" in vals else "DS"]


def test_period_ascending():
    ws = openpyxl.load_workbook(XLSX)["재무구조분석표"]
    heads = [ws.cell(4, c).value for c in range(2, 6)]
    assert heads == ["2023.12", "2024.12", "2025.12", "2026.09"], heads


def test_excel_sum_formulas():
    ws = openpyxl.load_workbook(XLSX)["재무구조분석표"]
    row = [r for r in range(5, 30) if ws.cell(r, 1).value == "소계(A) 순영업자산"][0]
    assert str(ws.cell(row, 2).value).startswith("=SUM(")


def test_word_landscape():
    sec = Document(DOCX).sections[0]
    assert sec.orientation != WD_ORIENT.LANDSCAPE  # 4개 기간 + 구분 = 5열 → 세로
    six = TEXT.replace("| 구분 | 2026.09.30 | 2025.12.31 | 2024.12.31 | 2023.12.31 |",
                       "| 구분 | 2026.09.30 | 2025.12.31 | 2024.12.31 | 2023.12.31 | 2022.12.31 |")
    six = "\n".join(l + " 0 |" if l.startswith("| ") and "구분" not in l and "---" not in l else l for l in six.splitlines())
    _, d2 = fst.build(six, "테스트", tempfile.mkdtemp(), OP)
    assert Document(d2).sections[0].orientation == WD_ORIENT.LANDSCAPE


def test_opinion_five_groups():
    heads = [p.text for p in Document(DOCX).paragraphs if p.style.name.startswith("Heading")]
    assert "확인 필요 사항" not in heads
    assert heads[-5:] == ["요지", "주요 발견사항", "규명된 핵심 사실", "리스크 요인", "후속 심사 과제(비공개 자료)"], heads
    bad = dict(OPINION, **{"확인 필요 사항": ["x"]})
    p = os.path.join(tempfile.mkdtemp(), "b.json"); json.dump(bad, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    try:
        fst.build(TEXT, "t", tempfile.mkdtemp(), p)
    except ValueError:
        return
    raise AssertionError("확인 필요 사항 묶음이 통과했다")


def test_delta_d_sum():
    periods, vals, _ = fst.parse_struct(TEXT)
    calc = [fst.compute(vals, j) for j in range(len(periods))]
    d = fst.delta_d(calc, periods.index("2025.12"), periods.index("2026.09"))
    assert d["D"] == 24 and abs(d["A"] + d["B"] + d["C"] + d["D"]) <= 2, d  # D = 단기성+장기성(엑셀 SUM과 같은 정의)


def test_checks_clean():
    mp = os.path.join(D, "m.json")
    json.dump({"company": "t", "unit": "억원", "bases": {}}, open(mp, "w", encoding="utf-8"))
    assert checks.run(DOCX, mp, "연결") == []


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
