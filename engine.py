"""숫자의 단일 출처: dart_std_data 결과 → 표준재무부표 v2.9 양식 채우기 → LibreOffice 재계산 → metrics.json·qa.json.
지표 산식은 양식 '주요 재무지표' 시트 수식이 유일한 기준이다. 여기서 산식을 다시 짜지 않는다.
사용: python3 kit/engine.py raw/std.txt --out data [--struct raw/struct.txt] [--pf raw/pf.txt]
"""
import argparse, json, os, re, shutil, subprocess, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, "template", "표준재무부표_템플릿_v2.9.xlsx")
TOL = 3.0  # 억원
import stdfs_fill  # noqa: E402  (같은 폴더)


def period_key(label: str) -> tuple:
    m = re.match(r"(\d{4})(?:\.(\d{2}))?", label.strip())
    return (int(m.group(1)), int(m.group(2) or 12))


def check_order(periods: list):
    """연간(12월)은 오름차순, 누적 기간은 연간 뒤에 오름차순. 어긋나면 AssertionError."""
    cum = [p for p in periods if "누적" in p]
    ann = [p for p in periods if "누적" not in p]
    assert periods == ann + cum, f"누적 기간이 연간 사이에 있다: {periods}"
    for g in (ann, cum):
        ks = [period_key(p) for p in g]
        assert ks == sorted(ks), f"기간이 오름차순이 아니다: {g}"


def clean(v):
    if isinstance(v, str):
        return None if not v.strip() else v
    return v


def recalc(xlsx: str) -> str:
    out = tempfile.mkdtemp()
    r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, xlsx],
                       capture_output=True, text=True, timeout=180)
    res = os.path.join(out, os.path.basename(xlsx))
    if r.returncode != 0 or not os.path.exists(res):
        raise RuntimeError(f"LibreOffice 재계산 실패: {r.stderr.strip()[:300]}")
    return res


def read_basis(wb, basis: str, n: int) -> dict:
    ws = wb[f"주요 재무지표({basis})"]
    periods = [ws.cell(4, 2 + j).value for j in range(n)]
    rows = {}
    for r in range(5, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if not isinstance(lab, str) or lab.startswith(("▶", "▷", "자료")):
            continue
        k = lab.strip()
        if k not in rows:  # 같은 이름(EBITDA 등)은 첫 행을 쓴다
            rows[k] = [clean(ws.cell(r, 2 + j).value) for j in range(n)]
    fs = wb[f"표준부표({basis})"]
    x = {}
    for r in range(1, fs.max_row + 1):
        c = fs.cell(r, 1).value
        if isinstance(c, str) and re.fullmatch(r"X[1-7]", c):
            x[c] = [clean(fs.cell(r, 3 + j).value) for j in range(n)]
    return {"periods": periods, "rows": rows}, x


def parse_struct_debt(text: str) -> dict:
    """구조표 결과의 총차입금 행 → {'YYYY.MM': 억원}."""
    unit = re.search(r"단위:\s*(원|천원|백만원|억원)", text)
    div = {"원": 1e8, "천원": 1e5, "백만원": 100, "억원": 1}[unit.group(1) if unit else "백만원"]
    head = next(l for l in text.splitlines() if l.startswith("| 구분"))
    periods = [c.strip()[:7] for c in head.strip("|").split("|")[1:]]
    row = next(l for l in text.splitlines() if re.match(r"\|\s*\**총차입금", l))
    vals = []
    for c in row.strip("|").split("|")[1:]:
        c = c.strip().replace(",", "")
        neg = c.startswith("(")
        v = float(c.strip("()")) if re.search(r"\d", c) else 0.0
        vals.append((-v if neg else v) / div)
    return dict(zip(periods, vals))


def _amt(s: str) -> float:
    s = s.replace(",", "").strip("* ")
    return 0.0 if s in ("", "-") else float(s)


def parse_pf(text: str) -> dict:
    """dart_credit_extras(kind="pf") 결과 → 양식 M코드(억원). 기준 기간은 머리줄의 보고서 기준 연월."""
    out = {"period": re.search(r"보고서\s*\((\d{4}\.\d{2})\)", text).group(1)}
    m = re.search(r"PF 신용보강\(브릿지론\+본PF\)\s*([\d,]+)", text)
    out["M13"] = _amt(m.group(1)) if m else 0.0
    bridge = main = 0.0
    for l in text.splitlines():  # 착공 판정 표의 판정별 행(보수적 기준 합계 행은 뺀다)
        c = [x.strip() for x in l.strip("|").split("|")]
        if len(c) == 5 and c[0] in ("미착공", "착공 미확인", "착공(지연 의심)", "착공"):
            bridge += _amt(c[1]); main += _amt(c[2])
    out["M17"], out["M18"] = bridge, main
    for code, pat in (("M14", r"책임준공 약정금액\s*([\d,]+)"), ("M15", r"중도금·이주비 보증\(회사분\)\s*([\d,]+)"),
                      ("M16", r"SOC 신용보강\(회사분\)\s*([\d,]+)")):
        m = re.search(pat, text)
        out[code] = _amt(m.group(1)) if m else 0.0
    return {k: (int(v) if isinstance(v, float) and v.is_integer() else v) for k, v in out.items()}


def fill_pf(xlsx: str, pf: dict) -> str | None:
    """연결 표준부표의 같은 기준 기간 열에 PF 금액을 넣는다. 맞는 열이 없으면 이유를 돌려준다."""
    wb = openpyxl.load_workbook(xlsx)
    ws = wb["표준부표(연결)"]
    row = {ws.cell(r, 1).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value}
    col = next((c for c in range(3, 10) if isinstance(ws.cell(4, c).value, str)
                and period_key(ws.cell(4, c).value) == period_key(pf["period"])), None)
    if col is None:
        return f"PF 명세 기준 기간 {pf['period']}이 재무 기간에 없어 양식에 넣지 않았다"
    for k, v in pf.items():
        if k.startswith("M") and k in row:
            ws.cell(row[k], col, v)
    wb.save(xlsx)
    return None


def run(std_path: str, out_dir: str, struct_path: str | None = None, pf_path: str | None = None) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    text = open(std_path, encoding="utf-8").read()
    head, data = stdfs_fill.parse(text)
    name = head.get("name", "")
    filled = os.path.join(tempfile.mkdtemp(), f"표준재무부표_{name}_v2.9.xlsx")
    stdfs_fill.fill(head, data, TPL, filled)
    pf_notes = []
    if pf_path:
        msg = fill_pf(filled, parse_pf(open(pf_path, encoding="utf-8").read()))
        if msg:
            pf_notes.append(msg)
    res = recalc(filled)
    final = os.path.join(out_dir, os.path.basename(filled))
    shutil.copy(res, final)
    wb = openpyxl.load_workbook(final, data_only=True)
    n = min(7, len(head["cols"]))
    bases, qa = {}, {"server_notes": {}, "x_checks": {}, "struct_diff": [], "pf_notes": pf_notes}
    for basis in ("연결", "별도"):
        if not data.get(basis, {}).get("codes"):
            continue
        b, x = read_basis(wb, basis, n)
        check_order(b["periods"])
        b["src"] = head["src"][:n]
        bases[basis] = b
        qa["server_notes"][basis] = data[basis]["notes"]
        qa["x_checks"][basis] = x
    metrics = {"company": name, "unit": "억원", "bases": bases}
    if struct_path and "연결" in bases:
        sd = parse_struct_debt(open(struct_path, encoding="utf-8").read())
        b = bases["연결"]
        for p, v in zip(b["periods"], b["rows"].get("총차입금", [])):
            s = next((sv for sp, sv in sd.items() if period_key(sp) == period_key(p)), None)
            if v is not None and s is not None and abs(v - s) > TOL:
                qa["struct_diff"].append({"period": p, "metrics_D": v, "struct_D": s, "diff": round(v - s, 1)})
    json.dump(metrics, open(os.path.join(out_dir, "metrics.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(qa, open(os.path.join(out_dir, "qa.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return metrics


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("std"); ap.add_argument("--out", default="data"); ap.add_argument("--struct")
    ap.add_argument("--pf", help="dart_credit_extras(kind='pf') 결과 원문 → 연결 '지급보증·신용보강' 행에 채운다")
    a = ap.parse_args()
    m = run(a.std, a.out, a.struct, a.pf)
    print("saved", a.out, {k: v["periods"][-1] for k, v in m["bases"].items()})
