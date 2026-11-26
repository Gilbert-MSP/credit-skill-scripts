"""dart_std_data 출력(텍스트) → 표준재무부표 v2.9 양식 엑셀.  사용: python stdfs_fill.py 데이터.txt 양식.xlsx 결과.xlsx"""
import re
import sys

import openpyxl


def parse(text):
    head, data, cur = {}, {}, None
    for ln in text.splitlines():
        ln = ln.strip()
        if ln.startswith("# "):
            m = re.search(r"·\s*(.+?)\s*\(억원", ln)
            head["name"] = m.group(1) if m else ""
        elif ln.startswith(("cols:", "months:", "src:")):
            k, v = ln.split(":", 1)
            head[k] = [x.strip() for x in v.split("|")]
        elif ln.startswith("## "):
            cur = ln[3:].strip()
            data[cur] = {"codes": {}, "notes": []}
        elif cur and re.match(r"^[A-Z]{1,2}\d{0,2}:", ln):
            k, v = ln.split(":", 1)
            data[cur]["codes"][k] = [float(x) if x.strip() else None for x in v.split(",")]
        elif cur and ln.startswith(("대사:", "확인할 것:")):
            data[cur]["notes"].append(ln)
    return head, data


def fill(head, data, tpl, out):
    wb = openpyxl.load_workbook(tpl)
    name = head.get("name", "")
    for basis in ("연결", "별도"):
        ws = wb[f"표준부표({basis})"]
        row = {ws.cell(r, 1).value: r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value}
        ws["A1"] = f"{name} 표준재무부표 — {basis} (K-IFRS, 단위: 억원)"
        wb[f"주요 재무지표({basis})"]["A1"] = f"{name} 주요 재무지표 — {basis} (K-IFRS, 단위: 억원)"
        codes = data.get(basis, {}).get("codes", {})
        if basis == "별도" and "R17" in codes:      # 별도 손익에는 지배·비지배 구분이 없다
            codes["R18"] = list(codes["R17"])
        for j, lab in enumerate(head["cols"][:7]):
            L = "CDEFGHI"[j]
            ws[f"{L}4"] = lab
            if not any(v[j] is not None for v in codes.values() if j < len(v)):
                continue
            ws[f"{L}{row['Z1']}"] = lab.replace("(누적)", "")
            ws[f"{L}{row['Z2']}"] = head["src"][j]
            ws[f"{L}{row['Z3']}"] = f"K-IFRS {basis}"
            ws[f"{L}{row['Z4']}"] = round(12 / int(head["months"][j]), 4)
            for k, vs in codes.items():
                if k in row and j < len(vs) and vs[j] is not None:
                    c = ws[f"{L}{row[k]}"]
                    if not (isinstance(c.value, str) and c.value.startswith("=")):   # 산식 칸은 건드리지 않는다
                        c.value = vs[j]
        r = ws.max_row + 2
        for t in data.get(basis, {}).get("notes", []) + ["자료: DART " + ", ".join(sorted({s for s in head["src"] if s and s != "-"}))]:
            ws.cell(r, 1, t)
            r += 1
    wb.save(out)


if __name__ == "__main__":
    h, dt = parse(open(sys.argv[1], encoding="utf-8").read())
    fill(h, dt, sys.argv[2], sys.argv[3])
    print("saved", sys.argv[3])
