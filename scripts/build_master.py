"""펀드 마스터 초안 생성 + 검증 큐 출력.

입력
  --site    사이트 공식 펀드 목록 JSON (/api/official-funds 응답)
  --model   Ver2_CF_모델.xlsx (자산 속성·T1 근거 텍스트)
출력 (--out 폴더)
  fund_master.json         확정 매칭 + 미확인 제안(verified=false)
  검증큐.csv               사람이 확인할 항목 (엑셀로 열기)

규칙
  - 엑셀 자산명이 사이트 공식명/별칭과 정규화 일치 → 자동 확정 별칭
  - 그 외 매칭 제안, 지분율 후보, 클래스(C/S/D) 제안은 전부 verified=false로 큐에 올린다
"""
from __future__ import annotations

import argparse
import collections
import csv
import json
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from yedaham.cf.excel_io import load_inputs  # noqa: E402
from yedaham.master.registry import Alias, Fund, Registry  # noqa: E402

RE_SHARE = re.compile(r"(당사|우리|지분)[^%]{0,30}?([0-9]{1,2}(?:\.[0-9]+)?|100(?:\.0+)?)\s*%")


def propose_class(a) -> tuple[str, str]:
    """Ver3 재설계 §4-2 기준의 1차 제안. 확정은 담당자."""
    if a.kind in ("파이프라인", "계획"):
        return "S", "신규 편입 — 운용사 페이싱(IM) 기반"
    if a.structure == "프로젝트":
        return "C", "프로젝트형 — 계약 결정"
    if a.region == "국내":
        return "S", "국내 블라인드 — 운용사 자금계획·공문"
    return "D", "해외 블라인드 — GP 재량"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--model", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    reg = Registry.from_site_official(args.site)
    inp = load_inputs(args.model)
    queue: list[dict] = []

    def q(kind, subject, proposal, basis, fid=""):
        queue.append({"구분": kind, "대상": subject, "제안": proposal, "근거": basis,
                      "펀드ID": fid, "확인(O/X)": "", "수정값": ""})

    # 1) 엑셀 CF 자산 ↔ 사이트 펀드
    share_cands: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    share_src: dict[tuple, str] = {}
    for e in inp.t1:
        for m in RE_SHARE.finditer(e.basis):
            v = float(m.group(2))
            share_cands[e.asset][v] += 1
            share_src.setdefault((e.asset, v), e.basis[:120])

    for a in inp.assets:
        m = reg.resolve(a.name, category="대체투자")
        if m.method == "exact":
            f = m.fund
            reg.add_alias(f, a.name, "excel_model", verified=True)
        elif a.kind == "계획":
            continue   # 계획은 펀드가 아니라 가정
        else:
            if m.method == "suggest" or (m.candidates and m.score >= 0.4):
                f = m.fund or m.candidates[0]
                reg.add_alias(f, a.name, "excel_model", verified=False)
                q("매칭", a.name, f.name, m.why, f.fid)
            else:
                f = Fund(f"new:{len(reg.funds) + 1}", a.name.replace("[신규] ", ""), "대체투자", "운용중",
                         [Alias(a.name, "excel_model", True)])
                reg.funds.append(f)
                reg._rebuild()
                q("신규등록", a.name, "사이트에 없는 펀드 — 신규 등록", "엑셀 파이프라인 자산", f.fid)
        # 속성
        f.ccy = f.ccy or a.ccy
        cls, why = propose_class(a)
        q("클래스", f.name, cls, why, f.fid)
        cands = share_cands.get(a.name)
        if cands:
            best = cands.most_common()
            q("지분율", f.name, " / ".join(f"{v}%({n}건)" for v, n in best),
              share_src[(a.name, best[0][0])], f.fid)
        else:
            q("지분율", f.name, "", "근거 없음 — 신탁·투자계약서 확인 필요", f.fid)

    # 2) 엑셀 채권 ↔ 사이트 채권
    for name in sorted({b.name for b in inp.bonds}):
        m = reg.resolve(name, category="채권")
        if m.method == "exact":
            reg.add_alias(m.fund, name, "excel_model", True)
        else:
            f = m.fund or (m.candidates[0] if m.candidates else None)
            if f:
                reg.add_alias(f, name, "excel_model", verified=False)
            q("매칭", name, f.name if f else "", m.why, f.fid if f else "")

    json.dump(reg.to_json(), open(os.path.join(args.out, "fund_master.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    with open(os.path.join(args.out, "검증큐.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(queue[0].keys()))
        w.writeheader()
        w.writerows(queue)
    c = collections.Counter(x["구분"] for x in queue)
    print(f"펀드 {len(reg.funds)}개 · 검증 항목 {len(queue)}건 {dict(c)}")


if __name__ == "__main__":
    main()
