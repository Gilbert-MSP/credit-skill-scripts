"""저장소 위생: 키 없음, 양식 시트 구성, 양식 채우기 동작. 실행: python3 tests/test_repo_clean.py"""
import os, re, subprocess, sys, tempfile
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(ROOT, "template", "표준재무부표_템플릿_v2.9.xlsx")
SAMPLE = os.path.join(ROOT, "tests", "fixtures", "std_sample.txt")


def test_no_secrets():
    bad = re.compile(r"(OPENDART_API_KEY\s*=|ACCESS_TOKEN\s*=|[A-Za-z0-9]{40,})")
    hits = []
    for d, dirs, files in os.walk(ROOT):
        dirs[:] = [x for x in dirs if x not in (".git", ".superpowers", "__pycache__")]
        for f in files:
            if f.endswith((".xlsx", ".pyc", ".png", ".docx")) or f == "test_repo_clean.py":
                continue
            p = os.path.join(d, f)
            if bad.search(open(p, encoding="utf-8", errors="ignore").read()):
                hits.append(p)
    assert not hits, hits


def test_template_sheets():
    assert openpyxl.load_workbook(TPL).sheetnames == [
        "주요 재무지표(연결)", "주요 재무지표(별도)", "표준부표(연결)", "표준부표(별도)",
        "PF 사업장 명세", "메자닌·PRS", "설정", "계정 매핑"]


def test_fill_runs():
    out = os.path.join(tempfile.mkdtemp(), "o.xlsx")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "stdfs_fill.py"), SAMPLE, TPL, out])
    assert r.returncode == 0 and os.path.exists(out)


if __name__ == "__main__":
    for k, f in list(globals().items()):
        if k.startswith("test_"):
            f(); print("ok", k)
