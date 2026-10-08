#!/usr/bin/env bash
# 엑셀 모델을 LibreOffice로 재계산한 사본을 만든다 (회귀 테스트의 정답값).
# 사용: scripts/recalc_excel.sh <원본.xlsx> <출력폴더>
set -euo pipefail
src="$1"; out="${2:-./_recalc}"
mkdir -p "$out"
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$out" "$src" >/dev/null
echo "$out/$(basename "$src")"
