"""약식 신용점검 Word 2~3쪽. 숫자는 metrics.json, 문장은 text.json(Claude 작성)에서만 읽는다.
사용: python3 kit/light_report.py data/metrics.json data/text.json <회사>_약식점검_YYYYMMDD.docx [연결|별도]
"""
import datetime, json, sys
import docx_kit as K
import checks

# (표시 이름, metrics 지표명, 서식)
TABLE1 = [("매출액", "매출액", "num"), ("EBITDA", "EBITDA", "num"), ("EBITDA 마진", "EBITDA/매출액", "pct"),
          ("총차입금", "총차입금", "num"), ("순차입금(총차입금−금융자산)", "순차입금(총차입금−금융자산)", "num"),
          ("순차입금/EBITDA", "순차입금/EBITDA(배)", "x"), ("EBITDA/금융비용", "EBITDA/금융비용(배)", "x"),
          ("부채비율", "부채비율", "pct"), ("운전자본 투자 후 영업현금", "운전자본 투자 후 영업현금", "num"),
          ("잉여현금흐름(FCF)", "잉여현금흐름(FCF)", "num")]
BOX = [("매출액", "num"), ("EBITDA", "num"), ("총차입금", "num"), ("순차입금/EBITDA(배)", "x"),
       ("EBITDA/금융비용(배)", "x"), ("부채비율", "pct")]
NOTE1 = ("주: 누적 기간의 손익·현금흐름은 해당 개월 누계이고 비율은 연환산이다. 총차입금은 리스부채를 포함하며 "
         "자본성 증권은 표준재무부표 '설정' 시트 기본값으로 반영했다. EBITDA = 매출액 − 매출원가 − 판관비 + 감가상각비 + 무형자산상각비이다.")


def _vals(b, key, idx, ebitda=None):
    row = b["rows"].get(key) or [None] * len(b["periods"])
    out = [row[i] for i in idx]
    if key == "순차입금/EBITDA(배)" and ebitda:
        out = ["EBITDA 적자" if (ebitda[i] is not None and ebitda[i] <= 0) else v for i, v in zip(idx, out)]
    return out


def build(metrics_path: str, text_path: str, out_path: str, basis: str | None = None) -> list:
    m = json.load(open(metrics_path, encoding="utf-8"))
    t = json.load(open(text_path, encoding="utf-8"))
    if len(t.get("issues", [])) > 3:
        raise ValueError("약식 점검의 쟁점은 3건 이하이다")
    basis = basis or ("연결" if "연결" in m["bases"] else "별도")
    b = m["bases"][basis]
    ps = b["periods"]
    src = K.source_line(m, basis)
    all_idx = list(range(len(ps)))
    ann = [i for i, p in enumerate(ps) if "누적" not in p]
    box_idx = ann[-2:] + [len(ps) - 1] if "누적" in ps[-1] else ann[-3:]
    eb = b["rows"].get("EBITDA")

    d = K.new_doc(f"{m['company']} 약식 신용점검 ({basis})")
    d.add_paragraph(f"기준: {ps[-1]} · 작성일 {datetime.date.today():%Y.%m.%d} · 단위 억원 · K-IFRS {basis} · 목적: 정기 모니터링")
    K.summary_box(d, t["conclusion"] + (f" 외부 신용등급은 {t['rating']}이다." if t.get("rating") else ""),
                  ["구분"] + [ps[i] for i in box_idx],
                  [[k] + _vals(b, k, box_idx, eb) for k, _ in BOX], src, [f for _, f in BOX])

    d.add_heading("1. 핵심 재무지표", 2)
    K.add_table(d, 1, "핵심 재무지표 추이", "억원, %, 배", ["구분"] + ps,
                [[name] + _vals(b, key, all_idx, eb) for name, key, _ in TABLE1], NOTE1, src, [f for *_, f in TABLE1])

    d.add_heading("2. 유동성·차입구조", 2)
    d.add_paragraph(t["liquidity"])
    li = ann[-1:] + ([len(ps) - 1] if "누적" in ps[-1] else [])
    K.add_table(d, 2, "유동성 지표", "배, %", ["구분"] + [ps[i] for i in li],
                [["현금성자산/단기성차입금"] + _vals(b, "현금성자산/단기성차입금", li),
                 ["단기성차입금/총차입금"] + _vals(b, "단기성차입금/총차입금", li)],
                "주: 현금성자산은 현금및현금성자산·단기금융상품·단기투자자산 합계이다.", src, ["x", "pct"])

    d.add_heading("3. 우발채무·특수관계자", 2)
    d.add_paragraph(t["contingent"])
    if t.get("pf"):
        d.add_paragraph(t["pf"])

    if t.get("issues"):
        d.add_heading("4. 쟁점 점검", 2)
        for i, x in enumerate(t["issues"], 1):
            p = d.add_paragraph()
            K._font(p.add_run(f"({i}) {x['title']}"), 9.5, True, K.NAVY)
            d.add_paragraph(f"공시 확인: {x['disclosure']}")
            d.add_paragraph(f"판단: {x['judgment']}")

    d.add_heading("5. 조기경보 지표", 2)
    K.add_table(d, 3, "조기경보 지표", "–", ["지표", "현재값", "경계 수준"],
                [[w["name"], w["now"], w["limit"]] for w in t["warnings"]],
                "주: 경계 수준은 공개 기준이 없으면 심사역 기준이다.", src)

    d.add_heading("6. 후속 심사 과제(비공개 자료)", 2)
    if t.get("followups"):
        for f in t["followups"]:
            d.add_paragraph(f"{f['what']} — {f['why']}. 판단 기준: {f['criterion']}.".replace("..", "."))
    else:
        d.add_paragraph("공시 외 자료로 확인할 과제는 없다.")
    d.save(out_path)
    return checks.run(out_path, metrics_path, basis)


if __name__ == "__main__":
    res = build(*sys.argv[1:4], sys.argv[4] if len(sys.argv) > 4 else None)
    print("saved", sys.argv[3])
    for x in res:
        print("-", x)
    sys.exit(1 if res else 0)
