"""CF 엔진 입력 모델.

엑셀 `Ver2_CF_모델.xlsx`의 Back Data 시트(01~08, 13)를 그대로 옮긴 구조다.
단위는 엑셀과 같이 억원(원화 환산은 현물 앵커 기준), 기간은 'YYYY-MM' 문자열.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Assumptions:
    """01_가정 — 전역 스위치."""
    opening_liquidity: float          # B4 기초 유동성
    fin_yield: float                  # B5 금융상품 수익률 (현재 엔진에서는 08의 요율을 씀)
    blind_draw_cap: float             # B6 블라인드 소진상한
    project_draw_cap: float           # B7 프로젝트 소진상한
    fx_scen_usd: float                # B8
    fx_scen_eur: float                # B9
    div_mult: float                   # B10 배당 시나리오 배율
    repay_mult: float                 # B11 상환 시나리오 배율
    call_mult: float                  # B12 콜 시나리오 배율
    plan_switch: float                # B13 신규 계획 편입 (1=계획선, 0=기준선)
    bond_wht: float                   # B14 채권 원천징수율 (표시용)
    liquidation_months: float         # B15 청산 판정 임계(개월) (상태 판정은 별도 단계)
    reinvest_decay: float             # B16 재투자 콜 감쇠율


@dataclass
class FXPath:
    """02_BD_환율 — 월별 원본 경로와 현물 앵커."""
    months: list[str]                 # 창 월 목록 (17개)
    usd: list[float]                  # USD 원본 경로
    eur: list[float]                  # EUR 원본 경로
    spot_usd: float                   # 현물 앵커
    spot_eur: float


@dataclass
class Curves:
    """03_BD_곡선 — 중분류 × 경과연수(Y0~Y12). 값이 없으면 None."""
    draw: dict[str, list[Optional[float]]]     # 인출률 (미인출 대비, 연율)
    runoff: dict[str, list[Optional[float]]]   # 상환률 (투입잔액 대비)
    dividend: dict[str, list[Optional[float]]] # 배당률 (투입잔액 대비) — 현재 엔진 미사용


@dataclass
class DivOverride:
    """연간 배당의 자산별 수동 지정 (엑셀에서 수식이 다른 행)."""
    kind: str                  # 'yield' = 투입잔액 × rate, 'fixed' = 고정값
    value: float
    note: str = ""


@dataclass
class Asset:
    """04_BD_자산마스터의 입력 열 + 05 캘린더 + 07 파이프라인 (같은 행)."""
    name: str                          # A
    kind: str                          # B 구분: 대체 | 파이프라인 | 계획
    structure: str                     # C 유형: 프로젝트 | 블라인드
    midclass: str                      # D 중분류
    strategy: str                      # E
    sector: str                        # F
    region: str                        # G
    ccy: str                           # H KRW | USD | EUR | AUD
    vintage: Optional[int]             # I
    age: Optional[float]               # J 경과연수
    commitment: float                  # K 약정
    invested_bal: float                # L 투입잔액
    unfunded: float                    # M 유효미인출
    cum_called: float                  # N 누적콜
    call_12m: float                    # P 직전12M 콜
    div_12m: float                     # Q 직전12M 배당
    div_24m_ann: float                 # R 직전24M 배당(연환산)
    repay_36m: float                   # S 직전36M 상환
    repay_rate_3y: float               # T 3년평균 상환율
    recurrence: Optional[str]          # U 경상성 등급
    sale_recovery: Optional[float]     # W 매각회수율
    call_basis: Optional[str]          # AP 콜 근거
    div_basis: Optional[str]           # AQ 배당 근거
    trust_end_months: Optional[float]  # AU 신탁종료 잔여(개월)
    rec_div_12m: float                 # AV 경상 12M 배당
    rec_div_24m: float                 # AW 경상 24M 배당
    status: str                        # AX Asset Status
    k3_invalid: Optional[float]        # BN
    k3_invalid_note: Optional[str]     # BO
    invest_end: Optional[str]          # BP 투자종료 (YYYY-MM-DD)
    confirmed_unfunded: Optional[float]  # BS 확정 미인출 (없으면 None)
    unfunded_note: Optional[str]       # BT
    win_invest_months: float           # BV 창내 기간중월
    win_post_months: float             # BW 창내 종료후월
    # 05_BD_캘린더 — 달력월(1~12) 플래그, 매각은 창 월(17개) 플래그
    cal_call: list[int] = field(default_factory=lambda: [0] * 12)
    cal_div: list[int] = field(default_factory=lambda: [0] * 12)
    cal_repay: list[int] = field(default_factory=lambda: [0] * 12)
    cal_sale: list[int] = field(default_factory=lambda: [0] * 17)
    # 07_BD_파이프라인
    pipe_pacing: list[float] = field(default_factory=lambda: [0.0] * 17)
    pipe_pacing_rule: Optional[dict] = None     # 잔여 페이싱 규칙 (계획 자산)
    pipe_div_start: Optional[str] = None        # S 배당개시
    pipe_div_rate: float = 0.0                  # T 배당률
    pipe_div_cal: list[int] = field(default_factory=lambda: [0] * 12)
    pipe_catchup: float = 0.0                   # AH 최초캐치업 (현재 엑셀 수식에서 미사용)
    # 자산별 수동 지정
    div_override: Optional[DivOverride] = None
    div_months_override: Optional[float] = None  # AG 배당월수 직접 입력


@dataclass
class T1Event:
    """06_BD_T1확정 — 확정 공문·실적·이동."""
    month: str
    asset: str
    kind: str          # 콜 | 배당 | 상환
    amount: float      # 억원 (당사분, 원화)
    etype: str         # 공문 | 실적 | 이동
    basis: str = ""


@dataclass
class FinProduct:
    """08_BD_금융상품 — 만기 이자 스케줄."""
    name: str
    maturity: str
    principal: float
    months_held: float
    interest: float
    tag: str


@dataclass
class FinParams:
    rate: float                 # B12 요율(연)
    scheduled_principal: float  # B13 이 금액까지는 만기 스케줄, 초과분은 월할


@dataclass
class BondLine:
    """13_채권별 — 한 줄 = 한 종목의 이자 또는 원금 스케줄."""
    name: str
    kind: str              # 이자 | 원금
    call_exercise: float   # 콜행사 스위치 (원금 줄에만 곱함)
    basis: str
    amounts: list[float]   # 창 17개월


@dataclass
class ModelInputs:
    assumptions: Assumptions
    fx: FXPath
    curves: Curves
    assets: list[Asset]
    t1: list[T1Event]
    fin_products: list[FinProduct]
    fin_params: FinParams
    bonds: list[BondLine]

    @property
    def months(self) -> list[str]:
        return self.fx.months
