"""엑셀 대비 회귀 테스트 — 엔진이 Ver2_CF_모델.xlsx와 같은 숫자를 내는지.

실데이터가 필요하므로 환경변수로 경로를 준다. 없으면 건너뛴다.
  YEDAHAM_MODEL_XLSX   : 수식 원본 (예: .../Cash Flow Ver.2/Ver2_CF_모델.xlsx)
  YEDAHAM_RECALC_XLSX  : 같은 파일을 LibreOffice로 재계산한 사본 (scripts/recalc_excel.sh)
"""
import os

import pytest

from yedaham.cf.engine import run
from yedaham.cf.excel_io import load_golden, load_inputs

SRC = os.environ.get("YEDAHAM_MODEL_XLSX")
RECALC = os.environ.get("YEDAHAM_RECALC_XLSX")
pytestmark = pytest.mark.skipif(not (SRC and RECALC), reason="실데이터 경로 미지정")
TOL = 1e-6   # 억원


@pytest.fixture(scope="module")
def both():
    return run(load_inputs(SRC)), load_golden(RECALC)


def test_every_asset_month_matches(both):
    res, g = both
    diffs = []
    for ar in res.assets:
        ga = g["assets"][ar.name]
        for key in ("call", "div", "repay"):
            for k, (x, y) in enumerate(zip(getattr(ar, key), ga[key])):
                if abs(x - y) > TOL:
                    diffs.append((ar.name, key, res.months[k], x, y))
    assert not diffs, f"{len(diffs)}건 불일치, 예: {diffs[:5]}"


def test_monthly_cf_matches(both):
    res, g = both
    for m, gm in zip(res.monthly, g["monthly"]):
        for f in ("call", "div", "repay", "bond_int", "bond_prin", "fin_int", "net", "closing"):
            assert abs(getattr(m, f) - gm[f]) < TOL, (m.month, f, getattr(m, f), gm[f])


def test_derived_budgets_match(both):
    res, g = both
    for ar in res.assets:
        gd = g["assets"][ar.name]["derived"]
        for f, v in gd.items():
            assert abs(getattr(ar.derived, f) - v) < TOL, (ar.name, f, getattr(ar.derived, f), v)
