# -*- coding: utf-8 -*-
"""2차시 '함수 중첩' 실습용/정답용 엑셀 파일 생성 스크립트.

입력
    문항_02_함수중첩.json     문항 정의 (답 셀, 정답 수식, 문제 문장)
    실습데이터_강좌운영.csv   [표1] 원본 (가상 자료)
    실습데이터_분류표.csv     [표2] 원본 (가상 자료)

출력
    실습_02_함수중첩.xlsx     답 셀을 비워 둔 실습 파일 (답 셀 연노랑 배경)
    정답_02_함수중첩.xlsx     답 셀에 정답 '수식'이 들어간 정답 파일

여러 번 실행해도 같은 결과가 나오며, 기존 파일은 덮어쓴다.

    python make_xlsx_02.py
"""

import csv
import json
import math
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

BASE = Path(__file__).resolve().parent
JSON_PATH = BASE / "문항_02_함수중첩.json"
CSV_MAIN = BASE / "실습데이터_강좌운영.csv"
CSV_CODE = BASE / "실습데이터_분류표.csv"
OUT_WORK = BASE / "실습_02_함수중첩.xlsx"
OUT_ANSWER = BASE / "정답_02_함수중첩.xlsx"

# --- 배치 상수 -----------------------------------------------------------
MAIN_ANCHOR = "A1"        # [표1] 왼쪽 위 (머리글 포함)
CODE_ANCHOR = "L1"        # [표2] 왼쪽 위 (머리글 포함)

# 문항 번호 -> 답 셀에 붙일 머리글/라벨.
# 범위형 답(G2:G15 등)은 1행 머리글로, 단일 셀 답(L10 등)은 바로 윗 셀에 쓴다.
ANSWER_LABEL = {
    1: "운영평가",
    2: "분류명",
    3: "과정구분",
    4: "[4번] 만족도 1위 강좌명",
    5: "[5번] IT 분류 평균 만족도",
}
# 답 셀 표시 형식 (지정하지 않으면 일반)
ANSWER_NUMFMT = {5: "0.00"}
# [표1] 숫자 열의 표시 형식
MAIN_NUMFMT = {"D": "0", "E": "0.0", "F": "0"}

# --- 서식 ---------------------------------------------------------------
FILL_HEADER = PatternFill("solid", fgColor="D9E2F3")   # 표 머리글: 연한 파랑
FILL_ANSWER = PatternFill("solid", fgColor="FFF2CC")   # 답 셀: 연한 노랑
FILL_NOTE = PatternFill("solid", fgColor="F2F2F2")     # 안내문: 연한 회색
FONT_HEADER = Font(bold=True)
FONT_TITLE = Font(bold=True, size=12)
_thin = Side(style="thin", color="A6A6A6")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
ALIGN_WRAP = Alignment(vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right")


def to_number(text):
    """숫자로 쓸 수 있는 문자열만 int/float로 바꾼다. 'O-101', 'IT'는 그대로 둔다."""
    s = text.strip()
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d*\.\d+", s):
        return float(s)
    return s


def read_csv(path):
    """CSV를 (머리글, 데이터 행 목록)으로 읽는다. 숫자 열은 숫자로 변환."""
    with path.open(encoding="utf-8", newline="") as f:
        rows = [r for r in csv.reader(f) if r]
    return rows[0], [[to_number(v) for v in r] for r in rows[1:]]


def parse_ref(ref):
    """'G2:G15' -> ('G', 2, 15),  'L10' -> ('L', 10, 10)"""
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", ref.strip())
    if not m:
        raise ValueError("답 셀 표기를 해석할 수 없습니다: %r" % ref)
    col, r1, col2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    if col2 and col2 != col:
        raise ValueError("답 셀은 한 열 안에 있어야 합니다: %r" % ref)
    return col, r1, int(r2) if r2 else r1


def anchor_cell(anchor):
    """'A1' -> (행 1, 열 1)"""
    m = re.fullmatch(r"([A-Z]+)(\d+)", anchor)
    return int(m.group(2)), column_index_from_string(m.group(1))


def write_table(ws, anchor, header, rows, numfmt=None):
    """머리글과 데이터를 anchor 위치부터 쓰고 테두리를 두른다."""
    r0, c0 = anchor_cell(anchor)
    for j, name in enumerate(header):
        cell = ws.cell(row=r0, column=c0 + j, value=name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER
        cell.alignment = ALIGN_CENTER
    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row):
            cell = ws.cell(row=r0 + i, column=c0 + j, value=value)
            cell.border = BORDER
            letter = get_column_letter(c0 + j)
            if numfmt and letter in numfmt:
                cell.number_format = numfmt[letter]
                cell.alignment = ALIGN_RIGHT


def build_sheet_data(ws, items, with_formula):
    """시트1 '강좌운영'을 만든다. with_formula가 참이면 답 셀에 정답 수식을 넣는다."""
    header, rows = read_csv(CSV_MAIN)
    write_table(ws, MAIN_ANCHOR, header, rows, MAIN_NUMFMT)
    code_header, code_rows = read_csv(CSV_CODE)
    write_table(ws, CODE_ANCHOR, code_header, code_rows)

    for item in items:
        no = item["번호"]
        col, r1, r2 = parse_ref(item["답셀"])
        label = ANSWER_LABEL[no]
        is_range = r2 > r1

        # 답 셀 위에 머리글(범위형) 또는 라벨(단일 셀)을 쓴다
        tag = ws["%s%d" % (col, r1 - 1)]
        tag.value = label
        tag.font = FONT_HEADER
        if is_range:
            tag.fill = FILL_HEADER
            tag.border = BORDER
            tag.alignment = ALIGN_CENTER

        base = item["정답수식"]
        origin = "%s%d" % (col, r1)
        for r in range(r1, r2 + 1):
            cell = ws["%s%d" % (col, r)]
            cell.fill = FILL_ANSWER
            cell.border = BORDER
            if no in ANSWER_NUMFMT:
                cell.number_format = ANSWER_NUMFMT[no]
            if with_formula:
                # 상대 참조만 행에 맞춰 옮기고 $L$2:$M$6 같은 절대 참조는 그대로 둔다
                target = "%s%d" % (col, r)
                cell.value = (base if r == r1 else
                              Translator(base, origin=origin)
                              .translate_formula(target))

    widths = (("A", 10), ("B", 24), ("C", 10), ("D", 10), ("E", 9), ("F", 9),
              ("G", 11), ("H", 12), ("I", 14), ("J", 3), ("K", 3),
              ("L", 24), ("M", 14))
    for letter, width in widths:
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"


def build_sheet_quiz(ws, data):
    """시트2 '문제'를 만든다."""
    ws["A1"] = "%d차시 %s - 실습 문제" % (data["차시"], data["차시명"])
    ws["A1"].font = FONT_TITLE
    ws["A2"] = ("시험영역 %s | 이 자료는 수업용 가상 자료입니다. "
                "'%s' 시트의 연노랑 셀에 답을 입력하세요."
                % (data["시험영역"], data["데이터"]["시트"]))
    ws["A2"].fill = FILL_NOTE
    ws.row_dimensions[2].height = 18

    header = ["번호", "문제", "답 셀", "사용 함수", "난이도"]
    start = 4
    for j, name in enumerate(header, start=1):
        cell = ws.cell(row=start, column=j, value=name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER
        cell.alignment = ALIGN_CENTER

    for i, item in enumerate(data["문항"], start=1):
        values = [item["번호"], item["문제"], item["답셀"],
                  ", ".join(item["사용함수"]), item["난이도"]]
        for j, value in enumerate(values, start=1):
            cell = ws.cell(row=start + i, column=j, value=value)
            cell.border = BORDER
            cell.alignment = ALIGN_CENTER if j in (1, 3, 5) else ALIGN_WRAP
        # 문제 문장 길이에 맞춰 행 높이를 잡는다 (한글은 한 줄에 약 45자)
        ws.row_dimensions[start + i].height = max(
            30, math.ceil(len(item["문제"]) / 45.0) * 16)

    for letter, width in (("A", 6), ("B", 92), ("C", 12), ("D", 26), ("E", 8)):
        ws.column_dimensions[letter].width = width


def build(data, path, with_formula):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = data["데이터"]["시트"]
    build_sheet_data(ws1, data["문항"], with_formula)
    build_sheet_quiz(wb.create_sheet("문제"), data)
    wb.active = 0
    wb.save(path)
    return path


def main():
    with JSON_PATH.open(encoding="utf-8") as f:
        data = json.load(f)

    build(data, OUT_WORK, with_formula=False)
    build(data, OUT_ANSWER, with_formula=True)

    print("생성 완료")
    print("  %s  (답 셀 비움)" % OUT_WORK.name)
    print("  %s  (답 셀에 수식)" % OUT_ANSWER.name)
    print("답 셀: " + ", ".join(
        "%d번 %s" % (i["번호"], i["답셀"]) for i in data["문항"]))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
