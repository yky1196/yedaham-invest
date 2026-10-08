# yedaham-invest

더케이예다함상조 대체투자 관리 시스템 (내부망 전용).

- CF 엔진: Cash Flow Ver.2 엑셀 모델의 python 이식 — 엑셀과 전 셀 일치 검증 (`docs/engine_spec.md`)
- 작업 원칙과 구조: `CLAUDE.md` · 결정 로그: `docs/decisions.md` · 로드맵: `docs/roadmap.md`

```bash
pip install -e ".[dev]"
python -m pytest -q
python -m yedaham.cf.cli "<데이터폴더>/Ver2_CF_모델.xlsx"   # 요약 출력
```

실데이터(원장·운용보고서·엑셀 모델)는 이 저장소에 넣지 않습니다.
