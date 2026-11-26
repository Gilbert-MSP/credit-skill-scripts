"""checks.py 테스트. 실행: python3 tests/test_checks.py"""
import json, os, sys, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import docx_kit as K  # noqa: E402
import checks  # noqa: E402

SRC = "자료: 사업보고서(2025.12), 연결"
M = {"company": "테스트", "unit": "억원", "bases": {"연결": {
    "periods": ["2024.12", "2025.12"], "src": ["사업보고서 (2025.12)"] * 2,
    "rows": {"매출액": [100, 120], "EBITDA": [20, 30], "총차입금": [50, 60]}}}}
_mp = os.path.join(tempfile.mkdtemp(), "metrics.json")
json.dump(M, open(_mp, "w", encoding="utf-8"), ensure_ascii=False)


def _doc(body=("결론은 양호한 수준이다.",), followups=(), table_debt=60, extra_table=False):
    d = K.new_doc("테스트 약식 점검 (연결)")
    for t in body:
        d.add_paragraph(t)
    K.add_table(d, 1, "핵심 재무지표", "억원", ["구분", "2024.12", "2025.12"],
                [["매출액", 100, 120], ["EBITDA", 20, 30], ["총차입금", 50, table_debt]], "주: 테스트", SRC)
    if extra_table:
        d.add_table(rows=1, cols=1).rows[0].cells[0].text = "출처 없는 표"
    d.add_heading("후속 심사 과제", 2)
    for f in followups:
        d.add_paragraph(f)
    p = os.path.join(tempfile.mkdtemp(), "t.docx"); d.save(p)
    return checks.run(p, _mp)


def _one(issues, word):
    assert len(issues) == 1 and word in issues[0], issues


def test_clean_doc_passes():
    assert _doc() == []


def test_ending_ham():
    _one(_doc(body=("차입 부담이 커짐.", "현금이 충분함. 다음 문장이다.")) [:1], "종결")
    assert len(_doc(body=("차입 부담이 커짐.", "현금이 충분함."))) == 2


def test_tag():
    _one(_doc(body=("매출이 늘었다[S01].",)), "태그")


def test_rcept():
    _one(_doc(body=("접수번호 20261112000704 공시이다.",)), "접수번호")


def test_author_phrase():
    _one(_doc(body=("작성자 계산 결과이다.",)), "작성자")


def test_source_count():
    _one(_doc(extra_table=True), "자료")


def test_followup_homework_blocked():
    _one(_doc(followups=("주석 24의 담보 내역을 확인해야 한다.",)), "후속")


def test_followup_private_allowed():
    assert _doc(followups=("약정서 원문으로 주석상 한도의 재무약정 조건을 확인한다.",)) == []


def test_number_mismatch():
    _one(_doc(table_debt=65), "총차입금")


def test_speculation_words():
    _one(_doc(body=("수요가 회복될 것으로 보인다.",)), "추론")


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
