# -*- coding: utf-8 -*-
"""낭독문에 기호·영문·셀 주소·괄호가 남아 있는지, 단계 큐가 순서대로 있는지 점검."""
import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
PATH = "대본_03_배열수식.md"
lines = open(PATH, encoding="utf-8").read().splitlines()

bad, cues = [], []
for i, line in enumerate(lines, 1):
    s = line.strip()
    if not s: continue
    if s.startswith("["):
        m = re.match(r"\[화면: 단계 (\d+)\]", s)
        if m: cues.append(int(m.group(1)))
        continue
    if s.startswith("#") or s.startswith(">"): continue
    hits = []
    if re.search(r"[A-Za-z]", s): hits.append("영문")
    if re.search(r"[=<>%*$:()\[\]{}]", s): hits.append("기호")
    if re.search(r"[A-Z]\d", s): hits.append("셀주소")
    if re.search(r"\d", s): hits.append("아라비아숫자")
    if hits: bad.append((i, ",".join(hits), s))

print("1) 낭독문 금지 표기:", "없음" if not bad else f"{len(bad)}건")
for i, h, s in bad: print(f"   {i}행 [{h}] {s}")
print("2) [화면: 단계 n] 큐:", cues, "→", "번호 순서 정상" if cues == list(range(1, len(cues)+1)) else "순서 이상")
print("3) 도입 첫 문장:", next(l for l in lines if l.startswith("안녕하세요")))
print("4) 구간 머리글:", [l[3:] for l in lines if l.startswith("## ")])
