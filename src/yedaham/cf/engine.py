"""CF 엔진 — Ver2_CF_모델.xlsx 수식의 python 이식.

원칙
- 1차 목표는 엑셀과 **같은 숫자**를 내는 것이다 (회귀 테스트로 강제).
- 엑셀 수식의 이상한 동작도 일단 그대로 재현하고, `docs/engine_spec.md`의
  '의문점' 목록에 남긴다. 고치는 것은 백테스트로 검증한 뒤 별도 결정으로 한다.

참조 표기: 각 계산 옆 주석의 (AM), (BK) 등은 04_BD_자산마스터의 열 이름이다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from .model import Asset, ModelInputs

KINDS = ("콜", "배당", "상환")


def month_of(ym: str) -> int:
    return int(ym[5:7])


# ---------------------------------------------------------------------------
# 곡선 조회 (X 인출률, AA 런오프율)
# ---------------------------------------------------------------------------
def curve_lookup(curve: dict[str, list[Optional[float]]], midclass: str, age: Optional[float]) -> float:
    """엑셀: IF(J>12,0, IFERROR(IF(COUNT(중분류값)>0, 중분류값, IFERROR(ALL값,0)),0))."""
    if age is None:
        return 0.0
    j = int(age)
    if j > 12 or j < 0:
        return 0.0
    row = curve.get(midclass)
    if row is not None and j < len(row) and row[j] is not None:
        return float(row[j])
    all_row = curve.get("ALL")
    if all_row is not None and j < len(all_row):
        return float(all_row[j]) if all_row[j] is not None else 0.0
    return 0.0


# ---------------------------------------------------------------------------
# 결과 구조
# ---------------------------------------------------------------------------
@dataclass
class AssetDerived:
    """04_BD_자산마스터의 수식 열."""
    draw_headroom: float = 0.0      # O 소진여력
    liquidating: bool = False       # V 청산
    draw_rate: float = 0.0          # X 인출률
    annual_call: float = 0.0        # Y
    annual_div: float = 0.0         # Z
    runoff_rate: float = 0.0        # AA
    repay_cand: tuple = (0.0, 0.0, 0.0)  # AB AC AD
    annual_repay: float = 0.0       # AE
    n_call_cal: int = 0             # AF
    n_div_cal: float = 0            # AG
    n_repay_cal: int = 0            # AH
    n_sale_win: int = 0             # AI
    t1_total: dict = field(default_factory=dict)       # AJ AK AL (전 유형)
    t1_doc_total: dict = field(default_factory=dict)   # BE BF BG ('실적' 제외)
    win_budget_call: float = 0.0    # AM 창내 콜예산
    win_budget_div: float = 0.0     # AN 창내 배당예산
    win_budget_repay: float = 0.0   # AO 창내 상환예산
    n_win: dict = field(default_factory=dict)          # AR AS AT 창내 횟수
    n_doc_months: dict = field(default_factory=dict)   # BH BI BJ 공문월수
    denom: dict = field(default_factory=dict)          # BK BL BM 분모
    reinvest_headroom: float = 0.0  # BR 재투자 여력
    annual_call_reinvest: float = 0.0  # BU 연간 콜(재투자)
    effective_unfunded: float = 0.0    # BS 있으면 BS, 없으면 M


@dataclass
class AssetResult:
    name: str
    ccy: str
    derived: AssetDerived
    call: list[float]
    div: list[float]
    repay: list[float]
    # 월별 근거 표시: 'T1' | 'PIPE' | 'MODEL' | 'SALE' | '0'
    call_src: list[str]
    div_src: list[str]
    repay_src: list[str]


@dataclass
class MonthRow:
    month: str
    opening: float
    call: float      # 음수
    div: float
    repay: float
    bond_int: float
    bond_prin: float
    fin_int: float
    net: float
    closing: float


@dataclass
class EngineResult:
    months: list[str]
    assets: list[AssetResult]
    monthly: list[MonthRow]

    @property
    def totals(self) -> dict:
        s = lambda f: sum(getattr(r, f) for r in self.monthly)  # noqa: E731
        trough = min(self.monthly, key=lambda r: r.closing)
        return {
            "call": s("call"), "div": s("div"), "repay": s("repay"),
            "bond_int": s("bond_int"), "bond_prin": s("bond_prin"), "fin_int": s("fin_int"),
            "net": s("net"), "trough": trough.closing, "trough_month": trough.month,
            "closing": self.monthly[-1].closing,
        }


# ---------------------------------------------------------------------------
# 엔진
# ---------------------------------------------------------------------------
class Engine:
    def __init__(self, inp: ModelInputs):
        self.inp = inp
        self.months = inp.months
        self.n = len(self.months)
        self.cal_idx = [month_of(m) - 1 for m in self.months]   # 창 k → 달력월 index(0~11)
        # T1 인덱스
        self.t1_month: dict[tuple, float] = {}     # (월, 자산, 구분) → 합계 (전 유형)
        self.t1_month_doc: dict[tuple, float] = {}  # '실적' 제외
        self.t1_all: dict[tuple, float] = {}        # (자산, 구분) → 합계 (전 유형)
        self.t1_all_doc: dict[tuple, float] = {}
        for e in inp.t1:
            k = (e.month, e.asset, e.kind)
            self.t1_month[k] = self.t1_month.get(k, 0.0) + e.amount
            self.t1_all[(e.asset, e.kind)] = self.t1_all.get((e.asset, e.kind), 0.0) + e.amount
            if e.etype != "실적":
                self.t1_month_doc[k] = self.t1_month_doc.get(k, 0.0) + e.amount
                self.t1_all_doc[(e.asset, e.kind)] = self.t1_all_doc.get((e.asset, e.kind), 0.0) + e.amount

    # ---- 환율 배율 (02_BD_환율 F~I열) ----
    def fx_ratio(self, k: int, ccy: str) -> float:
        a, fx = self.inp.assumptions, self.inp.fx
        usd = fx.usd[k] * a.fx_scen_usd / fx.spot_usd
        if ccy == "USD":
            return usd
        if ccy == "EUR":
            return fx.eur[k] * a.fx_scen_eur / fx.spot_eur
        if ccy == "AUD":   # 엑셀: I열 = F열 (AUD에 USD 배율 적용) — 의문점 Q3
            return usd
        return 1.0

    # ---- 04 자산마스터 수식 열 ----
    def derive(self, a: Asset) -> AssetDerived:
        A = self.inp.assumptions
        d = AssetDerived()
        cap = A.blind_draw_cap if a.structure == "블라인드" else A.project_draw_cap
        bs = a.confirmed_unfunded
        unf = bs if bs is not None else a.unfunded
        d.effective_unfunded = unf
        d.draw_headroom = max(a.commitment * cap - a.cum_called, 0.0)                     # O
        d.liquidating = a.status in ("Exit Planned", "Liquidating")                        # V
        d.draw_rate = curve_lookup(self.inp.curves.draw, a.midclass, a.age)                # X
        # Y 연간 콜
        hist = (a.k3_invalid != 1) and a.call_12m > 0.5
        y1 = a.call_12m if hist else unf * d.draw_rate
        y2 = bs if bs is not None else d.draw_headroom
        y3 = unf
        d.annual_call = min(y1, y2, y3) * A.call_mult
        # Z 연간 배당
        if a.div_override is not None:
            if a.div_override.kind == "yield":
                base = 0.0 if a.status == "Matured" else a.invested_bal * a.div_override.value
                d.annual_div = base * A.div_mult
            else:
                d.annual_div = a.div_override.value
        elif a.status == "Matured":
            d.annual_div = 0.0
        elif d.liquidating:
            d.annual_div = (a.rec_div_12m + a.rec_div_24m / 2) / 2 * A.div_mult
        else:
            d.annual_div = (a.div_12m + a.div_24m_ann) / 2 * A.div_mult
        # AA ~ AE 상환
        d.runoff_rate = curve_lookup(self.inp.curves.runoff, a.midclass, a.age)
        ab = a.repay_36m / 3
        ac = a.invested_bal * d.runoff_rate
        ad = a.invested_bal * a.repay_rate_3y
        d.repay_cand = (ab, ac, ad)
        d.n_call_cal = sum(a.cal_call)
        d.n_div_cal = a.div_months_override if a.div_months_override is not None else sum(a.cal_div)
        d.n_repay_cal = sum(a.cal_repay)
        d.n_sale_win = sum(a.cal_sale)
        if d.n_sale_win > 0 or a.status == "Matured":
            d.annual_repay = 0.0
        else:
            d.annual_repay = min(max(ab, ac, ad), a.invested_bal) * A.repay_mult
        # T1 합계
        for kind in KINDS:
            d.t1_total[kind] = self.t1_all.get((a.name, kind), 0.0)
            d.t1_doc_total[kind] = self.t1_all_doc.get((a.name, kind), 0.0)
        # 창내 횟수 / 공문월수 / 분모
        flags = {"콜": a.cal_call, "배당": a.cal_div, "상환": a.cal_repay}
        for kind in KINDS:
            f = flags[kind]
            d.n_win[kind] = sum(f[ci] for ci in self.cal_idx)
            d.n_doc_months[kind] = sum(
                1 for k, ci in enumerate(self.cal_idx)
                if f[ci] == 1 and self.t1_month_doc.get((self.months[k], a.name, kind), 0.0) > 0
            )
            d.denom[kind] = max(d.n_win[kind] - d.n_doc_months[kind], 1)
        # BR, BU
        d.reinvest_headroom = max(a.commitment * cap - a.invested_bal, 0.0)
        bu1 = a.call_12m * A.reinvest_decay if a.call_12m > 0.5 else 0.0
        d.annual_call_reinvest = min(bu1, max(d.reinvest_headroom, unf)) * A.call_mult
        # AM AN AO 창내 예산
        be, bf, bg = (d.t1_doc_total[k] for k in KINDS)
        d.win_budget_call = min(
            max(d.annual_call * a.win_invest_months / 12 + d.annual_call_reinvest * a.win_post_months / 12 - be, 0.0),
            max(max(unf, d.reinvest_headroom) - be, 0.0),
        )
        d.win_budget_div = max(d.annual_div * d.n_win["배당"] / max(d.n_div_cal, 1) - bf, 0.0)
        d.win_budget_repay = min(
            max(d.annual_repay * self.n / 12 - bg, 0.0),   # 엑셀 상수 1.4166… = 17/12
            max(a.invested_bal - bg, 0.0),
        )
        return d

    # ---- 월별 계산 ----
    def pacing(self, a: Asset) -> list[float]:
        """07 파이프라인 콜 페이싱. 계획 자산의 잔여 규칙(엑셀 P48:R48 수식)을 평가한다."""
        p = list(a.pipe_pacing)
        rule = a.pipe_pacing_rule
        if rule and rule.get("type") == "residual_curve":
            # 엑셀: IFERROR(INDEX(곡선, MATCH(중분류), year+1), 0) — ALL 폴백 없음
            row = self.inp.curves.draw.get(rule["midclass"]) or []
            y = rule.get("year", 0)
            rate = float(row[y]) if y < len(row) and row[y] is not None else 0.0
            for k in rule["months_idx"]:
                p[k] = max(0.0, (1 - sum(p[:k])) * rate / 12)
        return p

    def run_asset(self, a: Asset) -> AssetResult:
        A = self.inp.assumptions
        d = self.derive(a)
        months = self.months
        pace = self.pacing(a)
        call, div, rep = [0.0] * self.n, [0.0] * self.n, [0.0] * self.n
        csrc, dsrc, rsrc = ["0"] * self.n, ["0"] * self.n, ["0"] * self.n
        try:
            div_start_idx = months.index(a.pipe_div_start) if a.pipe_div_start else None
        except ValueError:
            div_start_idx = None
        is_pipe_div = a.kind == "파이프라인" or a.pipe_div_rate > 0
        pipe_div_n = max(sum(a.pipe_div_cal), 1)

        for k, ym in enumerate(months):
            ci = self.cal_idx[k]
            fx = self.fx_ratio(k, a.ccy)
            # --- 콜 (10_자산별_콜)
            t1 = self.t1_month.get((ym, a.name, "콜"), 0.0)
            if t1 > 0:
                call[k], csrc[k] = t1, "T1"
            else:
                if pace[k] > 0:
                    plan = A.plan_switch if a.kind == "계획" else 1.0
                    call[k], csrc[k] = a.commitment * pace[k] * plan * fx, "PIPE"
                elif a.cal_call[ci] == 1:
                    call[k], csrc[k] = d.win_budget_call / d.denom["콜"] * fx, "MODEL"
            # --- 배당 (11_자산별_배당)
            t1 = self.t1_month.get((ym, a.name, "배당"), 0.0)
            if t1 > 0:
                div[k], dsrc[k] = t1, "T1"
            elif is_pipe_div:
                if div_start_idx is not None and k >= div_start_idx and a.pipe_div_cal[ci] == 1:
                    # 엑셀: 누적 콜(이미 환율 반영된 원화) × 배당률 / 배당월수 × 환율 — 의문점 Q1
                    div[k] = sum(call[: k + 1]) * a.pipe_div_rate / pipe_div_n * fx
                    dsrc[k] = "PIPE"
            else:
                if div_start_idx is not None and k < div_start_idx:
                    pass
                elif a.cal_div[ci] == 1:
                    div[k], dsrc[k] = d.win_budget_div / d.denom["배당"] * fx, "MODEL"
            # --- 상환 (12_자산별_상환)
            t1 = self.t1_month.get((ym, a.name, "상환"), 0.0)
            if t1 > 0:
                rep[k], rsrc[k] = t1, "T1"
            elif a.cal_sale[k] == 1:
                rec = a.sale_recovery or 0.0
                rep[k], rsrc[k] = a.invested_bal * rec / max(d.n_sale_win, 1) * fx, "SALE"
            elif d.n_sale_win > 0:
                pass
            elif a.cal_repay[ci] == 1:
                rep[k], rsrc[k] = d.win_budget_repay / d.denom["상환"] * fx, "MODEL"
        return AssetResult(a.name, a.ccy, d, call, div, rep, csrc, dsrc, rsrc)

    def run(self) -> EngineResult:
        inp = self.inp
        assets = [self.run_asset(a) for a in inp.assets]
        # 채권
        bond_int = [0.0] * self.n
        bond_prin = [0.0] * self.n
        for b in inp.bonds:
            for k in range(self.n):
                if b.kind == "원금":
                    bond_prin[k] += round(b.amounts[k], 4) * b.call_exercise
                else:
                    bond_int[k] += b.amounts[k]
        # 금융상품 + 잔고 연결
        rows: list[MonthRow] = []
        opening = inp.assumptions.opening_liquidity
        fp = inp.fin_params
        for k, ym in enumerate(self.months):
            c = -sum(r.call[k] for r in assets)
            dv = sum(r.div[k] for r in assets)
            rp = sum(r.repay[k] for r in assets)
            sched = sum(f.interest for f in inp.fin_products if f.maturity == ym)
            fin = sched + max(opening - fp.scheduled_principal, 0.0) * fp.rate / 12
            net = c + dv + rp + bond_int[k] + bond_prin[k] + fin
            closing = opening + net
            rows.append(MonthRow(ym, opening, c, dv, rp, bond_int[k], bond_prin[k], fin, net, closing))
            opening = closing
        return EngineResult(self.months, assets, rows)


def run(inp: ModelInputs) -> EngineResult:
    return Engine(inp).run()
