"""출고 전 검사: 문체·태그·접수번호·출처 줄·숙제 문구·추론어·수치 대사.
사용: python3 kit/checks.py out.docx data/metrics.json [연결|별도]  → 문제가 있으면 종료 코드 1
"""
import json, re, sys
from docx import Document
from docx.oxml.ns import qn

ENDING = re.compile(r"(함|임|음|됨|짐|봄|옴|줌|림|김)\.(\s|$)")
TAG = re.compile(r"\[[SN]\d+\]")
RCEPT = re.compile(r"(?<!\d)\d{14}(?!\d)")
AUTHOR = re.compile(r"작성자 계산|분석자 설정")
SPEC = re.compile(r"보인다|아마|통상|듯하|듯이|할 듯|될 듯")
PUBLIC = re.compile(r"주석|공시|사업보고서|확인 필요|확인해야")
PRIVATE = re.compile(r"계약서|약정서|실사|면담")


def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _num(s: str):
    s = s.strip().replace(",", "").replace("−", "-")
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else None


def _body_items(doc):
    """(종류, 요소, 글) 목록. 종류: p | h | pic | tbl"""
    out = []
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            st = el.find(qn("w:pPr"))
            sty = st.find(qn("w:pStyle")).get(qn("w:val")) if st is not None and st.find(qn("w:pStyle")) is not None else ""
            kind = "h" if sty.startswith("Heading") else ("pic" if el.findall(".//" + qn("w:drawing")) else "p")
            out.append((kind, el, _text(el)))
        elif el.tag == qn("w:tbl"):
            out.append(("tbl", el, _text(el)))
    return out


def _sources(items) -> list:
    issues, missing = [], 0
    for i, (k, el, t) in enumerate(items):
        if k == "tbl":
            if "종합 심사 의견" in t:
                if "자료:" not in t:
                    missing += 1
                continue
        elif k != "pic":
            continue
        nxt = [x[2] for x in items[i + 1:i + 3] if x[0] == "p"]
        if not any(s.startswith("자료:") for s in nxt):
            missing += 1
    if missing:
        issues.append(f"자료: 줄이 없는 표·그림 {missing}개")
    for k, el, t in items:
        if t.startswith("자료:") and (not t[3:].strip() or re.search(r"TBD|…", t)):
            issues.append(f"자료: 줄이 비었거나 미완성이다: {t[:40]}")
    return issues


def _numbers(doc, metrics, basis) -> list:
    b = metrics["bases"].get(basis)
    if not b:
        return []
    issues = []
    for tbl in doc.element.body.iter(qn("w:tbl")):
        rows = tbl.findall(qn("w:tr"))
        if len(rows) < 2:
            continue
        head = [_text(c).strip() for c in rows[0].findall(qn("w:tc"))]
        cols = {j: b["periods"].index(h) for j, h in enumerate(head) if h in b["periods"]}
        if not cols:
            continue
        for r in rows[1:]:
            cells = [_text(c).strip() for c in r.findall(qn("w:tc"))]
            lab = cells[0] if cells else ""
            if lab not in b["rows"]:
                continue
            for j, pi in cols.items():
                if j >= len(cells):
                    continue
                v, ref = _num(cells[j]), b["rows"][lab][pi]
                if v is not None and isinstance(ref, (int, float)) and abs(v - ref) > 1:
                    issues.append(f"{lab} {head[j]} 표 값 {cells[j]}이 metrics.json {ref:,.0f}과 다르다")
    return issues


def run(docx_path: str, metrics_path: str, basis: str | None = None) -> list:
    doc = Document(docx_path)
    metrics = json.load(open(metrics_path, encoding="utf-8"))
    items = _body_items(doc)
    if basis is None:
        title = next((t for k, _, t in items if k == "h"), "")
        basis = "별도" if "별도" in title or "연결" not in metrics["bases"] else "연결"
    issues = []
    texts = [t for k, _, t in items if k in ("p", "h")]
    for tbl in doc.element.body.iter(qn("w:tbl")):
        texts += [_text(p) for p in tbl.iter(qn("w:p"))]
    for t in texts:
        for pat, msg in ((ENDING, "~함·~임·~음 종결"), (TAG, "대괄호 출처 태그"), (RCEPT, "접수번호"),
                         (AUTHOR, "작성자·분석자 문구"), (SPEC, "근거 없는 추론어")):
            if pat.search(t):
                issues.append(f"{msg}: {t[:50]}")
    issues += _sources(items)
    inside = False
    for k, _, t in items:
        if k == "h":
            inside = "후속" in t
        elif inside and k == "p" and PUBLIC.search(t) and not PRIVATE.search(t):
            issues.append(f"후속 과제에 공시로 확인할 항목이 남아 있다(직접 조회해 결론을 쓸 것): {t[:50]}")
    issues += _numbers(doc, metrics, basis)
    return issues


if __name__ == "__main__":
    res = run(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    for x in res:
        print("-", x)
    print("문제 없음" if not res else f"문제 {len(res)}건")
    sys.exit(1 if res else 0)
