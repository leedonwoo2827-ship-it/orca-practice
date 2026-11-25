# -*- coding: utf-8 -*-
"""대본의 낭독문 글자 수를 센다. 제목(#)·큐 줄([)·메모 줄(>)·빈 줄 제외, 공백 포함."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PATH = "대본_03_배열수식.md"
TARGET = {"도입": 300, "개념": 1100, "시연": 2200, "정리": 500}

def is_narration(line):
    s = line.strip()
    if not s: return False
    if s.startswith("#"): return False
    if s.startswith("["): return False
    if s.startswith(">"): return False
    return True

sections, cur = {}, None
for line in open(PATH, encoding="utf-8"):
    s = line.strip()
    if s.startswith("## "):
        cur = s[3:].strip(); sections[cur] = []
        continue
    if cur and is_narration(line):
        sections[cur].append(s)

total = 0
print(f"{'구간':<6}{'기준':>8}{'결과':>8}{'차이':>8}  문장")
for name, lines in sections.items():
    n = sum(len(l) for l in lines)
    total += n
    t = TARGET.get(name, 0)
    print(f"{name:<6}{t:>8}{n:>8}{n-t:>+8}  {len(lines)}")
print(f"{'합계':<6}{'4000~4500':>8}{total:>8}")
print("기준 충족" if 4000 <= total <= 4500 else "기준 벗어남 — 고쳐야 함")
