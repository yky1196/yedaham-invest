"""펀드 마스터 — 한 펀드를 부르는 여러 이름을 하나의 식별자로 묶는다.

매칭 단계 (사람 확인 원칙)
  1. exact   : 정규화한 이름이 공식명 또는 등록 별칭과 같다 → 자동 확정
  2. suggest : 고유 토큰(번호·영문 약칭·고유명사) 기준으로 후보가 하나뿐이다 → **사람 확인 후** 별칭 등록
  3. none / ambiguous : 확정 불가 → 검증 큐

확인된 매칭은 별칭으로 등록되어 다음부터 exact로 잡힌다. 이렇게 사람 공수가 시간이 갈수록 준다.
"""
from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Iterable, Optional

# 법적 형태·상투어 — 펀드를 구분하는 데 쓰이지 않는 말 (suggest 단계의 토큰 비교에서만 제거)
GENERIC = [
    "전문투자형", "일반사모", "사모", "특별자산", "부동산", "증권", "투자신탁", "투자회사", "신탁",
    "재간접형", "위탁관리", "모부동산", "주식회사", "합자회사", "사모투자", "투자", "전문", "일반",
    "제", "호", "종", "글로벌", "global", "인프라",
]
RE_PUNCT = re.compile(r"[\s\-_·.,()\[\]{}\"'/“”‘’]")


def normalize(s: str) -> str:
    """공백·괄호·구두점·법인표기를 지우고 소문자로. (기존 사이트 MCP의 _norm과 호환 + 전각 처리)"""
    s = unicodedata.normalize("NFKC", s or "")
    for t in ("㈜", "(주)", "주식회사"):
        s = s.replace(t, "")
    return RE_PUNCT.sub("", s).lower()


def tokens(s: str) -> set[str]:
    """식별 토큰: 숫자 묶음(예: '10-4', '2-3', '42'), 영문 단어, 한글 고유어(상투어 제거 후)."""
    s = unicodedata.normalize("NFKC", s or "").replace("㈜", " ")
    out: set[str] = set()
    for m in re.finditer(r"\d+(?:[-–]\d+)*", s):
        out.add(m.group(0).replace("–", "-"))
    for m in re.finditer(r"[A-Za-z][A-Za-z0-9]*", s):
        w = m.group(0).lower()
        if w not in ("global", "fund", "the", "of"):
            out.add(w)
    rest = re.sub(r"[A-Za-z0-9\s\-_·.,()\[\]{}\"'/“”‘’]", " ", s)
    for g in sorted(GENERIC, key=len, reverse=True):
        rest = rest.replace(g, " ")
    for w in rest.split():
        if len(w) >= 2:
            out.add(w)
    return out


@dataclass
class Alias:
    text: str
    source: str            # site_official | site_alias | excel_model | capcall_page | manual …
    verified: bool = True  # 사람이 확인했는가


@dataclass
class Fund:
    fid: str                           # 시스템 식별자 (사이트 id가 있으면 'site:<id>')
    name: str                          # 공식명
    category: str                      # 대체투자 | 채권
    status: str                        # 운용중 | 종료
    aliases: list[Alias] = field(default_factory=list)
    # 아래는 CF 엔진·모니터링에 필요한 마스터 속성 (출처 확인 전에는 None)
    our_share: Optional[float] = None  # 당사 지분율 (공문 금액 × 이 값 = 당사분)
    cf_class: Optional[str] = None     # C 계약형 | S 운용사 계획형 | D GP 재량형
    master_fund: Optional[str] = None  # 해외 마스터펀드명 (재간접)
    ccy: Optional[str] = None

    def names(self) -> list[str]:
        return [self.name] + [a.text for a in self.aliases]


@dataclass
class Match:
    query: str
    method: str                 # exact | suggest | ambiguous | none
    fund: Optional[Fund] = None
    candidates: list[Fund] = field(default_factory=list)
    score: float = 0.0
    why: str = ""


class Registry:
    def __init__(self, funds: Iterable[Fund] = ()):
        self.funds: list[Fund] = list(funds)
        self._rebuild()

    def _rebuild(self) -> None:
        self._exact: dict[str, list[Fund]] = {}
        for f in self.funds:
            for n in f.names():
                self._exact.setdefault(normalize(n), [])
                if f not in self._exact[normalize(n)]:
                    self._exact[normalize(n)].append(f)

    # ---- 적재 ----
    @classmethod
    def from_site_official(cls, path: str) -> "Registry":
        data = json.load(open(path, encoding="utf-8"))
        funds = []
        for d in data:
            f = Fund(f"site:{d['id']}", d["name"], d["category"], d.get("status") or "")
            f.aliases = [Alias(a, "site_alias") for a in d.get("aliases", [])]
            funds.append(f)
        return cls(funds)

    def add_alias(self, fund: Fund, text: str, source: str, verified: bool = True) -> None:
        if normalize(text) in {normalize(n) for n in fund.names()}:
            return
        fund.aliases.append(Alias(text, source, verified))
        self._rebuild()

    def by_id(self, fid: str) -> Fund:
        return next(f for f in self.funds if f.fid == fid)

    # ---- 매칭 ----
    def resolve(self, query: str, category: Optional[str] = None, active_only: bool = False) -> Match:
        pool = [f for f in self.funds
                if (category is None or f.category == category) and (not active_only or f.status == "운용중")]
        hit = [f for f in self._exact.get(normalize(query), []) if f in pool]
        if len(hit) == 1:
            return Match(query, "exact", hit[0], score=1.0, why="정규화 이름 일치")
        if len(hit) > 1:
            return Match(query, "ambiguous", candidates=hit, why="같은 이름이 여러 펀드에 등록됨")
        q = tokens(query)
        if not q:
            return Match(query, "none", why="식별 토큰 없음")
        scored = []
        for f in pool:
            best, why = 0.0, ""
            for n in f.names():
                t = tokens(n)
                if not t:
                    continue
                inter = q & t
                # 숫자 토큰이 양쪽에 있는데 다르면 다른 펀드 (예: 3호 vs 4호)
                qn, tn = {x for x in q if x[0].isdigit()}, {x for x in t if x[0].isdigit()}
                if qn and tn and not (qn & tn):
                    continue
                s = len(inter) / len(q | t)
                if s > best:
                    best, why = s, f"공통 토큰 {sorted(inter)}"
            if best > 0:
                scored.append((best, f, why))
        scored.sort(key=lambda x: -x[0])
        if not scored:
            return Match(query, "none", why="후보 없음")
        top = scored[0]
        second = scored[1][0] if len(scored) > 1 else 0.0
        if top[0] >= 0.5 and top[0] - second >= 0.2:
            return Match(query, "suggest", top[1], [s[1] for s in scored[:3]], top[0], top[2])
        return Match(query, "ambiguous", candidates=[s[1] for s in scored[:3]], score=top[0],
                     why=f"상위 후보 점수 차이 부족 ({top[0]:.2f} vs {second:.2f})")

    # ---- 저장 ----
    def to_json(self) -> list[dict]:
        return [{
            "fid": f.fid, "name": f.name, "category": f.category, "status": f.status,
            "aliases": [a.__dict__ for a in f.aliases], "our_share": f.our_share,
            "cf_class": f.cf_class, "master_fund": f.master_fund, "ccy": f.ccy,
        } for f in self.funds]

    @classmethod
    def from_json(cls, data: list[dict]) -> "Registry":
        funds = []
        for d in data:
            f = Fund(d["fid"], d["name"], d["category"], d["status"],
                     [Alias(**a) for a in d["aliases"]], d.get("our_share"), d.get("cf_class"),
                     d.get("master_fund"), d.get("ccy"))
            funds.append(f)
        return cls(funds)
