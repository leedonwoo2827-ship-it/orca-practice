# -*- coding: utf-8 -*-
"""3차시 실습·정답 엑셀 파일을 만들고, 정답 파일을 다시 열어 수식을 확인한다.

- 실습_03_배열수식.xlsx : 표 + 요약표 라벨, 답 칸(I2:I6)은 비움
- 정답_03_배열수식.xlsx : 같은 표 + 답 칸에 배열 수식(Ctrl+Shift+Enter 형태)
"""
import csv, io, json, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

SHEET = "판매실적"
PRACTICE = "실습_03_배열수식.xlsx"
ANSWER = "정답_03_배열수식.xlsx"

rows = list(csv.reader(open("데이터_03.csv", encoding="utf-8")))
header, data = rows[0], rows[1:]
assert len(data) == 20, f"데이터 행 수가 20이 아님: {len(data)}"

items = json.load(open("문항_03.json", encoding="utf-8"))["문항"]
LABELS = {
    1: '지점 "강북" + 상품군 "가전" 매출액 합계',
    2: '지점 "강남" + 분기 2 주문 건수',
    3: '지점 "수원" + 상품군 "주방" 판매수량 평균 (소수 1자리)',
    4: '상품군 "생활" 판매수량 최대값',
    5: '상품군 "주방" + 판매수량 20 이상 매출액 합계',
}
FMT = {3: "0.0"}

HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build(path, with_formula):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET

    # 표 머리글 (A1:F1)
    for c, name in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c, value=name if name != "매출액" else "매출액(만원)")
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BOX

    # 표 본문 (A2:F21) — 분기·판매수량·매출액은 숫자로 넣는다
    for r, row in enumerate(data, start=2):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=int(v) if c in (3, 5, 6) else v)
            cell.border = BOX
            if c in (3, 5, 6):
                cell.alignment = Alignment(horizontal="center" if c == 3 else "right")

    # 요약표 (H1:I6)
    for c, name in zip((8, 9), ("항목", "결과")):
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BOX
    for q in items:
        r = int(q["답셀"][1:])          # I2 -> 2
        ws.cell(row=r, column=8, value=f"{q['번호']}. {LABELS[q['번호']]}").border = BOX
        ans = ws.cell(row=r, column=9)
        ans.border = BOX
        ans.alignment = Alignment(horizontal="right")
        if q["번호"] in FMT:
            ans.number_format = FMT[q["번호"]]
        if with_formula:
            # 배열 수식(CSE). 엑셀에서 열면 수식 양 끝에 중괄호가 보인다.
            ans.value = ArrayFormula(q["답셀"], q["정답수식"])

    for col, w in zip("ABCDEF", (11, 22, 7, 10, 11, 14)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 44
    ws.column_dimensions["I"].width = 12
    ws.freeze_panes = "A2"

    wb.save(path)
    return path


print("=== 만들기 ===")
print(" ", build(PRACTICE, with_formula=False), "(답 칸 비움)")
print(" ", build(ANSWER, with_formula=True), "(답 칸에 배열 수식)")

print("\n=== 정답 파일 다시 열어 수식 확인 ===")
wb = load_workbook(ANSWER)
ws = wb[SHEET]
for q in items:
    v = ws[q["답셀"]].value
    text = v.text if isinstance(v, ArrayFormula) else v
    kind = "배열수식" if isinstance(v, ArrayFormula) else type(v).__name__
    ok = "일치" if text == q["정답수식"] else "다름!"
    print(f"  {q['답셀']}  [{kind}] {ok}  {text}")

print("\n=== 실습 파일 답 칸 비었는지 확인 ===")
ws2 = load_workbook(PRACTICE)[SHEET]
for q in items:
    print(f"  {q['답셀']} = {ws2[q['답셀']].value!r}")

print("\n=== 파이썬으로 따로 계산한 기대값 ===")
T = [dict(zip(header, r)) for r in data]
for r in T:
    r["분기"], r["판매수량"], r["매출액"] = int(r["분기"]), int(r["판매수량"]), int(r["매출액"])
sel = lambda p: [r for r in T if p(r)]
exp = {}
exp["I2"] = sum(r["매출액"] for r in sel(lambda r: r["지점"] == "강북" and r["상품군"] == "가전"))
exp["I3"] = len(sel(lambda r: r["지점"] == "강남" and r["분기"] == 2))
_s = [r["판매수량"] for r in sel(lambda r: r["지점"] == "수원" and r["상품군"] == "주방")]
exp["I4"] = round(sum(_s) / len(_s), 1)
exp["I5"] = max(r["판매수량"] for r in sel(lambda r: r["상품군"] == "생활"))
exp["I6"] = sum(r["매출액"] for r in sel(lambda r: r["상품군"] == "주방" and r["판매수량"] >= 20))
for k, v in exp.items():
    print(f"  {k} = {v}")
print("\n  (참고) 대본 단계 5 예외: 대전 + 가구 행 수 =",
      len(sel(lambda r: r["지점"] == "대전" and r["상품군"] == "가구")), "→ 합계 0, 평균은 #DIV/0!")
