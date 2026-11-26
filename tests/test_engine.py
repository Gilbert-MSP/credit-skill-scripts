"""engine.py 테스트. 실행: python3 tests/test_engine.py (LibreOffice 필요, 네트워크 없음)"""
import json, os, re, sys, tempfile
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import engine  # noqa: E402

FX = os.path.join(ROOT, "tests", "fixtures")
STD = os.path.join(FX, "std_sample.txt")
STRUCT = os.path.join(FX, "struct_sample.txt")
_cache = {}


def _run(std=STD, struct=None):
    key = (std, struct)
    if key not in _cache:
        d = tempfile.mkdtemp()
        _cache[key] = (engine.run(std, d, struct), d)
    return _cache[key]


def test_metrics_shape():
    m, d = _run()
    b = m["bases"]["연결"]
    assert len(b["periods"]) == 7
    assert b["periods"][0] == "2021.12" and b["periods"][-1] == "2026.09(누적)"
    for k in ("매출액", "EBITDA", "총차입금", "순차입금(총차입금−금융자산)", "부채비율"):
        assert k in b["rows"], k
    assert m["unit"] == "억원" and m["company"] == "한솔케미칼"
    assert json.load(open(os.path.join(d, "metrics.json"), encoding="utf-8")) == m


def test_matches_recalculated_xlsx():
    m, d = _run()
    x = [f for f in os.listdir(d) if f.endswith(".xlsx")][0]
    ws = openpyxl.load_workbook(os.path.join(d, x), data_only=True)["주요 재무지표(연결)"]
    assert ws.cell(30, 1).value == "총차입금"
    assert m["bases"]["연결"]["rows"]["총차입금"][-1] == ws.cell(30, 8).value == 2835


def test_blank_formula_is_none():
    assert engine.clean("") is None and engine.clean("  ") is None
    assert engine.clean(0) == 0 and engine.clean(1.5) == 1.5


def test_period_order_with_cumulative():
    # 5개년 오름차순 뒤에 전년동기·당기 누적이 오는 양식 순서를 지킨다
    ps = ["2021.12", "2022.12", "2023.12", "2024.12", "2025.12", "2025.09(누적)", "2026.09(누적)"]
    assert engine.check_order(ps) is None
    for bad in (["2022.12", "2021.12"], ["2025.12", "2026.09(누적)", "2025.09(누적)"]):
        try:
            engine.check_order(bad)
        except AssertionError:
            continue
        raise AssertionError(bad)
    assert engine.period_key("2026.09(누적)") == (2026, 9) and engine.period_key("2025") == (2025, 12)


def test_separate_only():
    t = open(STD, encoding="utf-8").read()
    head, rest = t.split("## 연결", 1)
    sep = "## 별도" + rest.split("## 별도", 1)[1]
    p = os.path.join(tempfile.mkdtemp(), "std.txt")
    open(p, "w", encoding="utf-8").write(head + sep)
    m, _ = _run(p)
    assert list(m["bases"]) == ["별도"]


def test_struct_diff():
    _, d = _run(struct=STRUCT)
    q = json.load(open(os.path.join(d, "qa.json"), encoding="utf-8"))
    assert [x["period"] for x in q["struct_diff"]] == ["2025.12", "2026.09(누적)"]  # 2,809 vs 2,805, 2,835 vs 2,830
    t = open(STRUCT, encoding="utf-8").read().replace("| **총차입금** | 2,830 | 2,805", "| **총차입금** | 2,830 | 3,805")
    p = os.path.join(tempfile.mkdtemp(), "s.txt")
    open(p, "w", encoding="utf-8").write(t)
    _, d2 = _run(struct=p)
    q2 = json.load(open(os.path.join(d2, "qa.json"), encoding="utf-8"))
    hit = [x for x in q2["struct_diff"] if x["period"] == "2025.12"]
    assert len(hit) == 1 and abs(hit[0]["diff"] + 996) < 1, hit  # metrics 2,809 − struct 3,805
    assert q2["server_notes"]["연결"] and "X1" in q2["x_checks"]["연결"]


def test_struct_unit_million():
    t = open(STRUCT, encoding="utf-8").read().replace("(단위: 억원)", "(단위: 백만원)").replace("| **총차입금** | 2,830 |", "| **총차입금** | 283,000 |")
    assert engine.parse_struct_debt(t)["2026.09"] == 2830


PF = os.path.join(FX, "pf_sample.txt")


def test_parse_pf():
    assert engine.parse_pf(open(PF, encoding="utf-8").read()) == {
        "period": "2026.09", "M13": 26432, "M17": 24877, "M18": 1556, "M14": 78318, "M15": 9645, "M16": 305}


def test_pf_fills_excel():
    d = tempfile.mkdtemp()
    m = engine.run(STD, d, pf_path=PF)
    r = m["bases"]["연결"]["rows"]
    assert r["지급보증·신용보강 제공액(연결 밖)"][-1] == 26432 + 78318 + 9645 + 305
    assert r["PF 신용보강(연대보증·채무인수·자금보충)"][-1] == 26432
    assert r["그중 브릿지론(본PF 전·미착공 단계)"][-1] == 24877
    assert r["그중 본PF"][-1] == 1556
    assert r["책임준공 조건부 채무인수"][-1] == 78318
    assert r["PF 신용보강(연대보증·채무인수·자금보충)"][-2] in (0, None)  # 다른 기간은 건드리지 않는다
    assert m["bases"]["별도"]["rows"]["PF 신용보강(연대보증·채무인수·자금보충)"][-1] in (0, None)


def test_pf_period_mismatch():
    t = open(PF, encoding="utf-8").read().replace("분기보고서 (2026.09)", "반기보고서 (2026.06)")
    p = os.path.join(tempfile.mkdtemp(), "pf.txt"); open(p, "w", encoding="utf-8").write(t)
    d = tempfile.mkdtemp()
    m = engine.run(STD, d, pf_path=p)
    assert m["bases"]["연결"]["rows"]["PF 신용보강(연대보증·채무인수·자금보충)"][-1] in (0, None)
    q = json.load(open(os.path.join(d, "qa.json"), encoding="utf-8"))
    assert any("PF" in x for x in q["pf_notes"])


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
