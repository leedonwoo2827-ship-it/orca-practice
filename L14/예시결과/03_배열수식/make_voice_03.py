# -*- coding: utf-8 -*-
"""대본 낭독문으로 낭독_03.txt(낭독 표기)와 자막_03.srt(화면 표기)를 만든다.

- 낭독본: 한 줄 한 문장. 규칙/발음사전_컴활.csv를 화면표기→낭독표기 방향으로 적용.
- 자막  : 함수 이름·셀 주소·기호를 화면 표기로 되돌림. 한 자막 최대 2줄, 한 줄 20자 안팎.
- 시각  : 초당 4.8자 (낭독문 글자 수 기준, 공백 포함).
"""
import csv, io, re, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SCRIPT = "대본_03_배열수식.md"
DICT = "../../규칙/발음사전_컴활.csv"
CPS = 4.8
LINE_LEN = 20

# ── 1. 대본에서 낭독문만 뽑아 한 줄 한 문장으로 ──────────────────
sents = []
for raw in open(SCRIPT, encoding="utf-8"):
    s = raw.strip()
    if not s or s.startswith(("#", "[", ">")):
        continue
    for part in re.split(r"(?<=\.)\s+", s):
        if part.strip():
            sents.append(part.strip())

# ── 2. 발음 사전을 화면표기 → 낭독표기 방향으로 적용 ─────────────
entries = [(r["화면표기"], r["낭독표기"]) for r in csv.DictReader(open(DICT, encoding="utf-8"))]
applied = []
narration = []
for s in sents:
    out = s
    for screen, read in sorted(entries, key=lambda e: -len(e[0])):
        if screen in out:
            out = out.replace(screen, read)
            applied.append(screen)
    narration.append(out)
open("낭독_03.txt", "w", encoding="utf-8").write("\n".join(narration) + "\n")

# ── 3. 자막용: 낭독 표기 → 화면 표기 되돌리기 ────────────────────
# 긴 것부터 바꾼다. 문맥까지 적은 항목은 다른 낱말을 건드리지 않게 하려는 것.
BACK = [
    ("이프에러", "IFERROR"), ("이프", "IF"),
    ("섬 함수", "SUM 함수"), ("섬 안에", "SUM 안에"), ("섬 괄호", "SUM 괄호"),
    ("맥스나 민", "MAX나 MIN"), ("맥스 자리에 민을", "MAX 자리에 MIN을"),
    ("맥스 함수", "MAX 함수"), ("맥스 안에", "MAX 안에"), ("맥스에", "MAX에"),
    ("맥스에는", "MAX에는"), ("맥스", "MAX"),
    ("아이이 셀", "I2 셀"), ("아이오 셀", "I5 셀"),
    ("브이비에이", "VBA"),
    ("컨트롤과 시프트", "Ctrl과 Shift"), ("컨트롤 시프트 엔터", "Ctrl+Shift+Enter"),
    ("엔터", "Enter"),
    ("삼천구백이십", "3,920"), ("천삼백육십", "1,360"), ("구백육십", "960"),
    ("천육백", "1,600"), ("오십오", "55"), ("사십팔", "48"), ("삼십오", "35"),
    ("사십", "40"),
    ("스무 개", "20개"), ("스무 건", "20건"),
    ("일 급", "1급"), ("삼 차시", "3차시"), ("사 차시", "4차시"),
]


def to_screen(t):
    for a, b in BACK:
        t = t.replace(a, b)
    return t


def wrap(t, width=LINE_LEN):
    """띄어쓰기를 지켜 width 안팎으로 자른다."""
    out, cur = [], ""
    for w in t.split(" "):
        if cur and len(cur) + 1 + len(w) > width:
            out.append(cur); cur = w
        else:
            cur = w if not cur else cur + " " + w
    if cur:
        out.append(cur)
    return out


def ts(sec):
    ms = int(round(sec * 1000))
    h, ms = divmod(ms, 3600000)
    m, ms = divmod(ms, 60000)
    s, ms = divmod(ms, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


cues, t, n = [], 0.0, 0
for read_sent in narration:
    dur = max(1.0, len(read_sent) / CPS)          # 낭독 글자 수 기준
    chunks = wrap(to_screen(read_sent))
    # 최대 두 줄씩 묶어 자막 한 개로
    blocks = [chunks[i:i + 2] for i in range(0, len(chunks), 2)]
    per = dur / len(blocks)
    for b in blocks:
        n += 1
        cues.append(f"{n}\n{ts(t)} --> {ts(t + per)}\n" + "\n".join(b) + "\n")
        t += per
open("자막_03.srt", "w", encoding="utf-8").write("\n".join(cues))

# ── 4. 보고용 수치 ────────────────────────────────────────────────
total_chars = sum(len(s) for s in narration)
over = [(i + 1, l) for i, c in enumerate(cues) for l in c.split("\n")[2:] if len(l) > 24]
print(f"낭독_03.txt  : {len(narration)}문장, {total_chars}자")
print(f"자막_03.srt  : 자막 {n}개, 전체 예상 길이 {int(t)//60}분 {int(t)%60}초 (초당 {CPS}자)")
print(f"사전 적용 항목: {sorted(set(applied)) or '없음'}")
print(f"한 줄 24자 초과 자막: {len(over)}개", over[:5] if over else "")
mx = max((len(l) for c in cues for l in c.split('\n')[2:]), default=0)
print(f"가장 긴 자막 줄: {mx}자")
