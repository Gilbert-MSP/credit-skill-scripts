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


def test_contract_derivative_rows():
    """계약자산·계약부채·파생상품 줄: 나눠 본 값의 합이 합친 줄과 같고, 파생상품부채 포함 총차입금 = 총차입금 + 파생상품부채(설정 B5 = 0)"""
    m, _ = _run()
    for b in m["bases"].values():
        r = b["rows"]
        for j in range(len(b["periods"])):
            ar, ca, tot = r["그중 매출채권"][j], r["그중 계약자산"][j], r["매출채권·계약자산"][j]
            if tot is not None:
                assert abs((ar or 0) + (ca or 0) - tot) < 0.01, (j, ar, ca, tot)
            d, q, ref = r["총차입금"][j], r["파생상품부채"][j], r["참고: 파생상품부채 포함 총차입금(전환권 등 상환 의무가 있을 때만 의미)"][j]
            if d is not None:
                assert abs(d + (q or 0) - ref) < 0.01, (j, d, q, ref)
        assert "계약부채(선수금 포함)" in r and "순수 운전자금(매출채권+재고−매입채무)" in r and "계약자산/매출액" in r


def test_rating_fin_cost():
    """평가사 기준 금융비용 = 이자비용 + 매출채권처분손실(R20) + 자본화 차입원가(R22, 설정 B12). 보증비용(R21)은 넣지 않는다"""
    base, _ = _run()
    txt = open(STD, encoding="utf-8").read().replace("## 연결\n", "## 연결\nR20:1,1,1,1,1,1,1\nR21:5,5,5,5,5,5,5\nR22:10,10,10,10,10,10,10\n", 1)
    p = os.path.join(tempfile.mkdtemp(), "std.txt")
    open(p, "w", encoding="utf-8").write(txt)
    m, _ = _run(p)
    k = "참고: 금융비용(평가사 기준: 이자비용+매출채권처분손실+자본화 이자(설정))"
    r0, r1 = base["bases"]["연결"]["rows"], m["bases"]["연결"]["rows"]
    assert r1[k][-1] == r0[k][-1] + 11, (r0[k], r1[k])
    assert r1["EBITDA/금융비용(배)"] == r0["EBITDA/금융비용(배)"]          # 기존 지표는 그대로
    assert r1["참고: EBITDA/금융비용(평가사 기준, 배)"][-1] < r0["참고: EBITDA/금융비용(평가사 기준, 배)"][-1]


def test_matches_recalculated_xlsx():
    m, d = _run()
    x = [f for f in os.listdir(d) if f.endswith(".xlsx")][0]
    ws = openpyxl.load_workbook(os.path.join(d, x), data_only=True)["주요 재무지표(연결)"]
    r = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, 1).value == "총차입금")   # 줄을 끼워 넣어도 이름으로 찾는다
    assert m["bases"]["연결"]["rows"]["총차입금"][-1] == ws.cell(r, 8).value == 2835


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
