"""엑셀 모델을 읽어 엔진을 돌리고 요약을 출력한다.

사용: python -m yedaham.cf.cli <Ver2_CF_모델.xlsx> [--golden <재계산본.xlsx>]
"""
from __future__ import annotations

import argparse

from .engine import run
from .excel_io import load_golden, load_inputs


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("model")
    ap.add_argument("--golden", help="LibreOffice 재계산본 — 주면 엑셀과 대조")
    args = ap.parse_args()
    res = run(load_inputs(args.model))
    t = res.totals
    print(f"창 {res.months[0]} ~ {res.months[-1]} ({len(res.months)}개월) · 억원")
    print(f"  Capital Call {t['call']:>10.2f}   배당 {t['div']:>8.2f}   상환 {t['repay']:>8.2f}")
    print(f"  채권 {t['bond_int'] + t['bond_prin']:>8.2f}   금융이자 {t['fin_int']:>6.2f}   순CF {t['net']:>8.2f}")
    print(f"  최저 버퍼 {t['trough']:.2f} ({t['trough_month']})   기말 {t['closing']:.2f}")
    print("\n월      기초     콜      배당    상환    채권    금융    기말")
    for m in res.monthly:
        print(f"{m.month} {m.opening:7.1f} {m.call:7.1f} {m.div:7.1f} {m.repay:7.1f} "
              f"{m.bond_int + m.bond_prin:7.1f} {m.fin_int:6.1f} {m.closing:7.1f}")
    if args.golden:
        g = load_golden(args.golden)
        n = sum(1 for ar in res.assets for key in ("call", "div", "repay")
                for x, y in zip(getattr(ar, key), g["assets"][ar.name][key]) if abs(x - y) > 1e-6)
        print(f"\n엑셀 대조: 불일치 셀 {n}건")


if __name__ == "__main__":
    main()
