"""엔진 규칙 단위 테스트 — 실데이터 없이 합성 입력으로 엑셀 수식의 의미를 고정한다."""
from yedaham.cf.engine import Engine, curve_lookup, run
from yedaham.cf.model import (Asset, Assumptions, BondLine, Curves, FinParams, FXPath,
                              ModelInputs, T1Event)

MONTHS = ["2026-08", "2026-09", "2026-10", "2026-11", "2026-12", "2027-01", "2027-02",
          "2027-03", "2027-04", "2027-05", "2027-06", "2027-07", "2027-08", "2027-09",
          "2027-10", "2027-11", "2027-12"]


def assum(**kw):
    base = dict(opening_liquidity=500, fin_yield=0.03, blind_draw_cap=0.95, project_draw_cap=1.0,
                fx_scen_usd=1, fx_scen_eur=1, div_mult=1, repay_mult=1, call_mult=1, plan_switch=1,
                bond_wht=0.154, liquidation_months=18, reinvest_decay=0.4)
    base.update(kw)
    return Assumptions(**base)


def asset(**kw):
    base = dict(name="A", kind="대체", structure="블라인드", midclass="코어", strategy="", sector="",
                region="", ccy="KRW", vintage=2022, age=2, commitment=100, invested_bal=60,
                unfunded=40, cum_called=60, call_12m=0, div_12m=0, div_24m_ann=0, repay_36m=0,
                repay_rate_3y=0, recurrence=None, sale_recovery=None, call_basis=None,
                div_basis=None, trust_end_months=None, rec_div_12m=0, rec_div_24m=0,
                status="Active", k3_invalid=None, k3_invalid_note=None, invest_end=None,
                confirmed_unfunded=None, unfunded_note=None, win_invest_months=17,
                win_post_months=0)
    base.update(kw)
    return Asset(**base)


def inputs(assets, t1=(), bonds=(), **akw):
    fx = FXPath(MONTHS, [1400.0] * 17, [1600.0] * 17, spot_usd=1400.0, spot_eur=1600.0)
    curves = Curves(draw={"ALL": [0.3] * 13, "코어": [0.5, 0.25, 0.4] + [None] * 10},
                    runoff={"ALL": [None, 0.03] + [0.05] * 11}, dividend={})
    return ModelInputs(assum(**akw), fx, curves, list(assets), list(t1), [],
                       FinParams(rate=0.031, scheduled_principal=400), list(bonds))


def test_curve_lookup_fallback_to_all():
    curves = {"ALL": [0.3] * 13, "코어": [0.5, None] + [None] * 11}
    assert curve_lookup(curves, "코어", 0) == 0.5
    assert curve_lookup(curves, "코어", 1) == 0.3       # 빈칸 → ALL
    assert curve_lookup(curves, "없는분류", 1) == 0.3   # 분류 없음 → ALL
    assert curve_lookup(curves, "코어", 13) == 0.0      # 경과 12년 초과 → 0


def test_call_budget_spread_over_calendar_months():
    # 분기 콜(3·6·9·12월), 직전12M 콜 없음 → 유효미인출 × 인출률(코어 Y2=0.4)
    a = asset(cal_call=[0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1])
    res = run(inputs([a]))
    r = res.assets[0]
    assert abs(r.derived.annual_call - 40 * 0.4) < 1e-12
    n_win = r.derived.n_win["콜"]          # 26-09, 26-12, 27-03, 27-06, 27-09, 27-12 = 6
    assert n_win == 6
    expected = 16 * 17 / 12 / 6
    assert all(abs(r.call[k] - expected) < 1e-9 for k in range(17) if r.call_src[k] == "MODEL")
    # 창내 예산은 유효미인출을 넘지 않는다 (창 단위 캡)
    assert sum(r.call) <= 40 + 1e-9


def test_t1_overrides_month_and_reduces_remaining_budget():
    a = asset(cal_call=[0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1])
    base = run(inputs([a])).assets[0]
    t1 = [T1Event("2026-09", "A", "콜", 5.0, "공문")]
    r = run(inputs([a], t1)).assets[0]
    assert r.call[1] == 5.0 and r.call_src[1] == "T1"
    # 공문 금액만큼 예산이 줄고 나머지 5개월에 재배분 → 창 합계 보존 (캡 미도달 시)
    assert abs(sum(r.call) - sum(base.call)) < 1e-9


def test_actual_t1_does_not_reduce_budget():
    """'실적' 유형은 그 달만 대체하고 예산·분모에는 영향이 없다 (엑셀 BE·BH의 <>실적 조건)."""
    a = asset(cal_call=[0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1])
    base = run(inputs([a])).assets[0]
    r = run(inputs([a], [T1Event("2026-09", "A", "콜", 5.0, "실적")])).assets[0]
    assert r.call[1] == 5.0
    assert abs(r.call[4] - base.call[4]) < 1e-12


def test_fx_ratio_applies_to_model_not_t1():
    a = asset(ccy="USD", cal_call=[1] * 12)
    inp = inputs([a], [T1Event("2026-08", "A", "콜", 3.0, "공문")])
    inp.fx.usd[1] = 1540.0          # 26-09 환율 10% 상승
    r = run(inp).assets[0]
    assert r.call[0] == 3.0         # T1은 이미 원화 확정값
    assert abs(r.call[1] / r.call[2] - 1.1) < 1e-12


def test_sale_calendar_replaces_runoff():
    a = asset(cal_repay=[1] * 12, cal_sale=[0] * 14 + [1, 0, 0], sale_recovery=0.8)
    r = run(inputs([a])).assets[0]
    assert r.derived.annual_repay == 0.0
    assert abs(r.repay[14] - 60 * 0.8) < 1e-12
    assert sum(r.repay) == r.repay[14]


def test_liquidating_uses_recurring_dividend():
    a = asset(status="Liquidating", div_12m=30, div_24m_ann=20, rec_div_12m=2, rec_div_24m=4,
              cal_div=[0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0])
    r = run(inputs([a])).assets[0]
    assert abs(r.derived.annual_div - (2 + 4 / 2) / 2) < 1e-12


def test_pipeline_pacing_and_plan_switch():
    a = asset(name="P", kind="계획", commitment=140, invested_bal=0, unfunded=140, cum_called=0,
              pipe_pacing=[0.0] * 13 + [0.3, 0, 0, 0])
    on = run(inputs([a])).assets[0]
    off = run(inputs([a], plan_switch=0)).assets[0]
    assert abs(on.call[13] - 140 * 0.3) < 1e-12
    assert off.call[13] == 0.0


def test_bond_principal_switch_and_cash_roll():
    bonds = [BondLine("B", "이자", 1, "", [1.0] * 17),
             BondLine("B", "원금", 0, "콜 미행사", [0.0] * 3 + [30.0] + [0.0] * 13)]
    res = run(inputs([asset()], bonds=bonds))
    assert sum(m.bond_prin for m in res.monthly) == 0.0
    assert abs(sum(m.bond_int for m in res.monthly) - 17) < 1e-12
    for prev, cur in zip(res.monthly, res.monthly[1:]):
        assert abs(cur.opening - prev.closing) < 1e-12
