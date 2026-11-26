"""숫자의 단일 출처: dart_std_data 결과 → 표준재무부표 v2.9 양식 채우기 → LibreOffice 재계산 → metrics.json·qa.json.
지표 산식은 양식 '주요 재무지표' 시트 수식이 유일한 기준이다. 여기서 산식을 다시 짜지 않는다.
사용: python3 kit/engine.py raw/std.txt --out data [--struct raw/struct.txt]
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


def run(std_path: str, out_dir: str, struct_path: str | None = None) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    text = open(std_path, encoding="utf-8").read()
    head, data = stdfs_fill.parse(text)
    name = head.get("name", "")
    filled = os.path.join(tempfile.mkdtemp(), f"표준재무부표_{name}_v2.9.xlsx")
    stdfs_fill.fill(head, data, TPL, filled)
    res = recalc(filled)
    final = os.path.join(out_dir, os.path.basename(filled))
    shutil.copy(res, final)
    wb = openpyxl.load_workbook(final, data_only=True)
    n = min(7, len(head["cols"]))
    bases, qa = {}, {"server_notes": {}, "x_checks": {}, "struct_diff": []}
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
    a = ap.parse_args()
    m = run(a.std, a.out, a.struct)
    print("saved", a.out, {k: v["periods"][-1] for k, v in m["bases"].items()})
